/// <reference path="../node_modules/@pixso/plugin-typings/index.d.ts" />

// Selection Variants & Instances — Pixso Plugin (Sandbox)
// This runs in the plugin sandbox with access to the Pixso API.

// ─── Types ───────────────────────────────────────────────

interface VariantProperty {
  name: string;
  currentValue: string;
  options: string[];
}

interface SwapOption {
  id: string;
  name: string;
}

interface ComponentPropertyInfo {
  name: string;
  type: string; // "BOOLEAN" | "TEXT" | "INSTANCE_SWAP" | "VARIANT"
  currentValue: string | boolean;
  currentValueName?: string; // resolved name for INSTANCE_SWAP
  currentValueSource?: string; // source/page info for INSTANCE_SWAP
  defaultValue?: string | boolean;
  preferredValues?: string[];
  swapOptions?: SwapOption[]; // resolved preferred values for INSTANCE_SWAP
  options?: string[]; // for VARIANT type
}

interface InstanceInfo {
  id: string;
  name: string;
  componentName: string;
  componentId: string | null;
  depth: number; // nesting level
  visible: boolean;
  variantProperties: VariantProperty[];
  componentProperties: ComponentPropertyInfo[];
  path: string; // parent chain for display
}

interface SelectionData {
  instances: InstanceInfo[];
  groupedByComponent: { [componentName: string]: InstanceInfo[] };
  totalSelected: number;
  hasSelection: boolean;
}

// ─── Helpers ─────────────────────────────────────────────

function isInstanceNode(node: SceneNode): node is InstanceNode {
  return node.type === "INSTANCE";
}

function isComponentNode(node: SceneNode): node is ComponentNode {
  return node.type === "COMPONENT";
}

function isComponentSetNode(node: SceneNode): node is ComponentSetNode {
  return node.type === "COMPONENT_SET";
}

function hasChildren(node: SceneNode): node is SceneNode & { children: readonly SceneNode[] } {
  return "children" in node;
}

function isEffectivelyVisible(node: SceneNode): boolean {
  let current: BaseNode | null = node;
  while (current && current.type !== "PAGE" && current.type !== "DOCUMENT") {
    if ("visible" in current && !(current as SceneNode).visible) {
      return false;
    }
    current = current.parent;
  }
  return true;
}

function getComponentDisplayName(instance: InstanceNode): string {
  const mainComponent = instance.mainComponent;
  if (!mainComponent) return "Unknown Component";

  // If mainComponent is inside a ComponentSet, use the set's name (master component name)
  const parent = mainComponent.parent;
  if (parent && parent.type === "COMPONENT_SET") {
    return parent.name;
  }

  // Otherwise use the component's own name
  return mainComponent.name;
}

function getNodePath(node: SceneNode): string {
  const parts: string[] = [];
  let current: BaseNode | null = node.parent;
  while (current && current.type !== "PAGE" && current.type !== "DOCUMENT") {
    parts.unshift(current.name);
    current = current.parent;
  }
  return parts.join(" > ");
}

function getVariantProperties(instance: InstanceNode): VariantProperty[] {
  const props: VariantProperty[] = [];
  const mainComponent = instance.mainComponent;
  if (!mainComponent) return props;

  const componentSet = mainComponent.parent;
  if (!componentSet || componentSet.type !== "COMPONENT_SET") return props;

  const variantGroupProps = (componentSet as ComponentSetNode).variantGroupProperties;
  if (!variantGroupProps) return props;

  const currentVariantProps = mainComponent.variantProperties;

  for (const [propName, propDef] of Object.entries(variantGroupProps)) {
    props.push({
      name: propName,
      currentValue: currentVariantProps?.[propName] ?? "",
      options: propDef.values,
    });
  }

  return props;
}

// Collect property names in the order they appear in the node tree (top-down)
// This matches Pixso's display order in the properties panel
function getPropertyOrderFromTree(node: SceneNode): string[] {
  const ordered: string[] = [];
  const seen = new Set<string>();

  function walk(n: SceneNode) {
    if ("componentPropertyReferences" in n && (n as any).componentPropertyReferences) {
      const refs = (n as any).componentPropertyReferences as { [key: string]: string };
      // visible (boolean) comes before mainComponent (swap) and characters (text)
      for (const refProp of ["visible", "characters", "mainComponent"]) {
        const propName = refs[refProp];
        if (propName && !seen.has(propName)) {
          seen.add(propName);
          ordered.push(propName);
        }
      }
    }
    if (hasChildren(n)) {
      for (const child of n.children) {
        walk(child);
      }
    }
  }

  walk(node);
  return ordered;
}

function getComponentProperties(instance: InstanceNode): ComponentPropertyInfo[] {
  const props: ComponentPropertyInfo[] = [];

  const compProps = instance.componentProperties;
  if (!compProps) return props;

  const mainComponent = instance.mainComponent;
  if (!mainComponent) return props;

  // Get definitions from the component or component set
  let definitions: ComponentPropertyDefinitions | null = null;
  const parent = mainComponent.parent;
  if (parent && parent.type === "COMPONENT_SET") {
    definitions = (parent as ComponentSetNode).componentPropertyDefinitions;
  } else {
    definitions = mainComponent.componentPropertyDefinitions;
  }

  // Get property order from the node tree (matches Pixso's display order)
  const treeOrder = getPropertyOrderFromTree(instance);

  // Start with tree-ordered properties, then append any remaining from definitions
  const defKeys = definitions ? Object.keys(definitions) : [];
  const orderedKeys: string[] = [];
  const added = new Set<string>();

  for (const propName of treeOrder) {
    if (definitions?.[propName] && compProps[propName]) {
      orderedKeys.push(propName);
      added.add(propName);
    }
  }
  // Append remaining (variant properties and any not found in tree)
  for (const propName of defKeys) {
    if (!added.has(propName) && compProps[propName]) {
      orderedKeys.push(propName);
    }
  }

  for (const propName of orderedKeys) {
    const propValue = compProps[propName];
    if (!propValue) continue;
    const def = definitions![propName];
    if (!def) continue;
    const info: ComponentPropertyInfo = {
      name: propName,
      type: propValue.type,
      currentValue: propValue.value,
      defaultValue: def.defaultValue,
    };

    if (propValue.type === "VARIANT" && def) {
      info.options = (def as any).variantOptions ?? [];
    }

    if (propValue.type === "INSTANCE_SWAP" && typeof propValue.value === "string") {
      // Find the child node controlled by this swap property
      // Walk all descendants looking for componentPropertyReferences.mainComponent match
      function findSwapChild(parent: SceneNode): SceneNode | null {
        if (!hasChildren(parent)) return null;
        for (const child of parent.children) {
          if ("componentPropertyReferences" in child) {
            const refs = (child as any).componentPropertyReferences;
            if (refs && refs.mainComponent === propName) {
              return child;
            }
          }
          const deeper = findSwapChild(child);
          if (deeper) return deeper;
        }
        return null;
      }

      const swapChild = findSwapChild(instance);

      if (swapChild) {
        // Layer name — what designer sees
        info.currentValueName = swapChild.name;

        if (isInstanceNode(swapChild)) {
          const mc = swapChild.mainComponent;
          if (mc) {
            // All metadata available directly on ComponentNode
            const mcP = mc.parent;
            const compName = (mcP && mcP.type === "COMPONENT_SET") ? mcP.name : mc.name;
            info.currentValueSource = compName;
            // Store full metadata for picker navigation
            (info as any).swapComponentKey = mc.key;
            (info as any).swapComponentName = compName;
            (info as any).swapPageName = (mc as any).pageName || "";
            (info as any).swapContainerName = (mc as any).containerName || "";
            (info as any).swapRemote = mc.remote;
          }
        }
      } else {
        const swapNode = pixso.getNodeById(propValue.value);
        if (swapNode) {
          info.currentValueName = swapNode.name;
          if ("key" in swapNode) {
            (info as any).swapComponentKey = (swapNode as any).key;
            (info as any).swapComponentName = swapNode.name;
            (info as any).swapPageName = (swapNode as any).pageName || "";
            (info as any).swapContainerName = (swapNode as any).containerName || "";
            (info as any).swapRemote = (swapNode as any).remote || false;
          }
        }
      }

      if (!info.currentValueName) {
        info.currentValueName = propName.replace(/#\d+:\d+$/, '').replace(/^[└─\s]+/, '').trim();
      }

      // Collect preferred values keys for lazy resolution
      if (def && "preferredValues" in def && (def as any).preferredValues) {
        info.preferredValues = (def as any).preferredValues.map(
          (pv: { type: string; key: string }) => pv.key
        );
      }
    } else if (def && "preferredValues" in def) {
      info.preferredValues = (def as any).preferredValues;
    }

    props.push(info);
  }

  return props;
}

function groupByComponent(instances: InstanceInfo[]): { [key: string]: InstanceInfo[] } {
  const grouped: { [key: string]: InstanceInfo[] } = {};
  for (const inst of instances) {
    const key = inst.componentName;
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(inst);
  }
  return grouped;
}

// Synchronous version for quick updates (property changes)
function analyzeSelectionSync(): SelectionData {
  const selection = pixso.currentPage.selection;
  if (!selection || selection.length === 0) {
    return { instances: [], groupedByComponent: {}, totalSelected: 0, hasSelection: false };
  }

  const results: InstanceInfo[] = [];
  const stack: { node: SceneNode; depth: number }[] = [];

  // Seed stack
  for (let i = selection.length - 1; i >= 0; i--) {
    stack.push({ node: selection[i], depth: 0 });
  }

  while (stack.length > 0) {
    const { node, depth } = stack.pop()!;

    if (isInstanceNode(node)) {
      const mainComponent = node.mainComponent;
      results.push({
        id: node.id,
        name: node.name,
        componentName: getComponentDisplayName(node),
        componentId: mainComponent?.id ?? null,
        depth,
        visible: isEffectivelyVisible(node),
        variantProperties: getVariantProperties(node),
        componentProperties: getComponentProperties(node),
        path: getNodePath(node),
      });
      if (hasChildren(node)) {
        for (let i = node.children.length - 1; i >= 0; i--) {
          stack.push({ node: node.children[i], depth: depth + 1 });
        }
      }
    } else if (hasChildren(node)) {
      for (let i = node.children.length - 1; i >= 0; i--) {
        stack.push({ node: node.children[i], depth: depth + 1 });
      }
    }
  }

  return {
    instances: results,
    groupedByComponent: groupByComponent(results),
    totalSelected: selection.length,
    hasSelection: true,
  };
}

// Async version with progress reporting
let analyzeAbortFlag = 0;

function analyzeSelectionAsync(onDone: (data: SelectionData) => void) {
  const selection = pixso.currentPage.selection;
  if (!selection || selection.length === 0) {
    onDone({ instances: [], groupedByComponent: {}, totalSelected: 0, hasSelection: false });
    return;
  }

  const myFlag = ++analyzeAbortFlag;
  const results: InstanceInfo[] = [];
  const stack: { node: SceneNode; depth: number }[] = [];

  for (let i = selection.length - 1; i >= 0; i--) {
    stack.push({ node: selection[i], depth: 0 });
  }

  let scanned = 0;
  const CHUNK_SIZE = 50;

  function processChunk() {
    if (myFlag !== analyzeAbortFlag) return; // aborted

    let processed = 0;
    while (stack.length > 0 && processed < CHUNK_SIZE) {
      const { node, depth } = stack.pop()!;
      scanned++;
      processed++;

      if (isInstanceNode(node)) {
        const mainComponent = node.mainComponent;
        results.push({
          id: node.id,
          name: node.name,
          componentName: getComponentDisplayName(node),
          componentId: mainComponent?.id ?? null,
          depth,
          visible: isEffectivelyVisible(node),
          variantProperties: getVariantProperties(node),
          componentProperties: getComponentProperties(node),
          path: getNodePath(node),
        });
        if (hasChildren(node)) {
          for (let i = node.children.length - 1; i >= 0; i--) {
            stack.push({ node: node.children[i], depth: depth + 1 });
          }
        }
      } else if (hasChildren(node)) {
        for (let i = node.children.length - 1; i >= 0; i--) {
          stack.push({ node: node.children[i], depth: depth + 1 });
        }
      }
    }

    // Send progress
    pixso.ui.postMessage({
      type: "scan-progress",
      scanned,
      found: results.length,
    });

    if (stack.length > 0) {
      setTimeout(processChunk, 0);
    } else {
      // Done
      onDone({
        instances: results,
        groupedByComponent: groupByComponent(results),
        totalSelected: selection.length,
        hasSelection: true,
      });
    }
  }

  setTimeout(processChunk, 0);
}

// ─── Find the variant ComponentNode matching desired property values ───

function findVariantComponent(
  instance: InstanceInfo,
  propertyName: string,
  newValue: string
): ComponentNode | null {
  // Get current instance node
  const instanceNode = pixso.getNodeById(instance.id) as InstanceNode | null;
  if (!instanceNode || instanceNode.type !== "INSTANCE") return null;

  const mainComponent = instanceNode.mainComponent;
  if (!mainComponent) return null;

  const componentSet = mainComponent.parent;
  if (!componentSet || componentSet.type !== "COMPONENT_SET") return null;

  // Build target variant properties
  const targetProps: { [key: string]: string } = {};
  for (const vp of instance.variantProperties) {
    targetProps[vp.name] = vp.name === propertyName ? newValue : vp.currentValue;
  }

  // Search through all variants in the component set
  for (const child of (componentSet as ComponentSetNode).children) {
    if (!isComponentNode(child)) continue;
    const variantProps = child.variantProperties;
    if (!variantProps) continue;

    let match = true;
    for (const [key, val] of Object.entries(targetProps)) {
      if (variantProps[key] !== val) {
        match = false;
        break;
      }
    }
    if (match) return child;
  }

  return null;
}

// ─── Message handling ────────────────────────────────────

console.log("[Selection Variants] Plugin starting...");
pixso.showUI(__html__, { width: 360, height: 520 });
// Enable window resize
try {
  (pixso.ui as any).enableResize = true;
  (pixso.ui as any).minWidth = 280;
  (pixso.ui as any).minHeight = 300;
} catch {}
console.log("[Selection Variants] UI shown");

// Check if current selection has a parent (can go up)
function getParentInfo(): { id: string; name: string } | null {
  const sel = pixso.currentPage.selection;
  if (!sel || sel.length === 0) return null;

  const parent = sel[0].parent;
  if (!parent || parent.type === "PAGE" || parent.type === "DOCUMENT") return null;

  return { id: parent.id, name: parent.name };
}

// Resolve INSTANCE_SWAP preferred values async, then send update to UI
async function resolveSwapOptions(data: SelectionData) {
  const keysToResolve = new Map<string, string>(); // key -> name (cache)

  for (const inst of data.instances) {
    for (const cp of inst.componentProperties) {
      if (cp.type === "INSTANCE_SWAP" && cp.preferredValues && cp.preferredValues.length > 0) {
        for (const key of cp.preferredValues) {
          if (!keysToResolve.has(key)) {
            keysToResolve.set(key, "");
          }
        }
      }
    }
  }

  if (keysToResolve.size === 0) return;

  // Resolve all keys in parallel
  const entries = Array.from(keysToResolve.keys());
  const resolved = await Promise.all(
    entries.map(async (key) => {
      try {
        const comp = await pixso.importComponentByKeyAsync(key);
        return { key, id: comp.id, name: comp.name };
      } catch {
        return null;
      }
    })
  );

  const keyMap = new Map<string, { id: string; name: string }>();
  for (const r of resolved) {
    if (r) keyMap.set(r.key, { id: r.id, name: r.name });
  }

  // Build swap options map: propertyName -> SwapOption[]
  const swapOptionsMap: { [propName: string]: SwapOption[] } = {};
  for (const inst of data.instances) {
    for (const cp of inst.componentProperties) {
      if (cp.type === "INSTANCE_SWAP" && cp.preferredValues && !swapOptionsMap[cp.name]) {
        const options: SwapOption[] = [];
        for (const key of cp.preferredValues) {
          const info = keyMap.get(key);
          if (info) options.push(info);
        }
        if (options.length > 0) {
          swapOptionsMap[cp.name] = options;
        }
      }
    }
  }

  if (Object.keys(swapOptionsMap).length > 0) {
    pixso.ui.postMessage({ type: "swap-options", swapOptionsMap });
  }
}

// Resolve unresolved INSTANCE_SWAP names by trying to find the node
// through parent instance's children
function resolveSwapNames(data: SelectionData) {
  for (const inst of data.instances) {
    for (const cp of inst.componentProperties) {
      if (cp.type === "INSTANCE_SWAP" && !cp.currentValueName && typeof cp.currentValue === "string") {
        // Find child whose componentPropertyReferences.mainComponent matches this property
        const instNode = pixso.getNodeById(inst.id) as InstanceNode | null;
        if (instNode && hasChildren(instNode)) {
          function findSwapChild(parent: SceneNode): boolean {
            if (!hasChildren(parent)) return false;
            for (const child of parent.children) {
              if ("componentPropertyReferences" in child) {
                const refs = (child as any).componentPropertyReferences;
                if (refs && refs.mainComponent === cp.name) {
                  if (isInstanceNode(child)) {
                    const mc = child.mainComponent;
                    if (mc) {
                      const mcParent = mc.parent;
                      if (mcParent && mcParent.type === "COMPONENT_SET") {
                        cp.currentValueName = mcParent.name + " / " + mc.name;
                      } else {
                        cp.currentValueName = mc.name;
                      }
                      return true;
                    }
                  } else if (child.name) {
                    cp.currentValueName = child.name;
                    return true;
                  }
                }
              }
              if (findSwapChild(child)) return true;
            }
            return false;
          }
          findSwapChild(instNode);
        }
      }
    }
  }
}

// Send data immediately (used after property changes — sync, no loading)
function sendSelectionData() {
  const data = analyzeSelectionSync();
  resolveSwapNames(data);
  const parentInfo = getParentInfo();
  pixso.ui.postMessage({
    type: "selection-data",
    data,
    parentInfo,
  });
}

// Send with loading + live progress (used on selection change)
function sendLoadingThenData() {
  pixso.ui.postMessage({ type: "loading" });

  analyzeSelectionAsync((data) => {
    const parentInfo = getParentInfo();
    pixso.ui.postMessage({
      type: "selection-data",
      data,
      parentInfo,
    });
  });
}

sendLoadingThenData();

// Listen for selection changes
pixso.on("selectionchange", () => {
  sendLoadingThenData();
});

// Listen for messages from UI
pixso.ui.on("message", (msg: any) => {
  if (msg.type === "set-variant") {
    // Switch a variant property on a specific instance
    const { instanceId, propertyName, newValue } = msg;
    const instanceNode = pixso.getNodeById(instanceId) as InstanceNode | null;
    if (!instanceNode || instanceNode.type !== "INSTANCE") return;

    // Find the instance info to get current variant props
    const data = analyzeSelectionSync();
    const instInfo = data.instances.find((i: InstanceInfo) => i.id === instanceId);
    if (!instInfo) return;

    const targetComponent = findVariantComponent(instInfo, propertyName, newValue);
    if (targetComponent) {
      const wasVisible = instanceNode.visible;
      instanceNode.swapComponent(targetComponent);
      instanceNode.visible = wasVisible;
    }

    sendSelectionData();
  }

  if (msg.type === "set-component-property") {
    // Set a component property (boolean, text, instance-swap)
    const { instanceId, propertyName, newValue } = msg;
    const instanceNode = pixso.getNodeById(instanceId) as InstanceNode | null;
    if (!instanceNode || instanceNode.type !== "INSTANCE") return;

    instanceNode.setProperties({ [propertyName]: newValue });
    sendSelectionData();
  }

  if (msg.type === "bulk-set-variant") {
    // Set variant property on ALL instances of the same component
    const { componentName, propertyName, newValue } = msg;
    const data = analyzeSelectionSync();
    const targetInstances = data.groupedByComponent[componentName] ?? [];

    for (const instInfo of targetInstances) {
      const instanceNode = pixso.getNodeById(instInfo.id) as InstanceNode | null;
      if (!instanceNode || instanceNode.type !== "INSTANCE") continue;

      const targetComponent = findVariantComponent(instInfo, propertyName, newValue);
      if (targetComponent) {
        const wasVisible = instanceNode.visible;
        instanceNode.swapComponent(targetComponent);
        instanceNode.visible = wasVisible;
      }
    }

    sendSelectionData();
  }

  if (msg.type === "bulk-set-component-property") {
    const { componentName, propertyName, newValue } = msg;
    const data = analyzeSelectionSync();
    const targetInstances = data.groupedByComponent[componentName] ?? [];

    for (const instInfo of targetInstances) {
      const instanceNode = pixso.getNodeById(instInfo.id) as InstanceNode | null;
      if (!instanceNode || instanceNode.type !== "INSTANCE") continue;

      instanceNode.setProperties({ [propertyName]: newValue });
    }

    sendSelectionData();
  }

  if (msg.type === "select-instance") {
    const { instanceId } = msg;
    const node = pixso.getNodeById(instanceId);
    if (node && "type" in node) {
      pixso.currentPage.selection = [node as SceneNode];
    }
  }

  if (msg.type === "select-instances-by-component") {
    const { componentName } = msg;
    const data = analyzeSelectionSync();
    const targetInstances = data.groupedByComponent[componentName] ?? [];
    const nodes: SceneNode[] = [];
    for (const inst of targetInstances) {
      const node = pixso.getNodeById(inst.id);
      if (node && "type" in node) {
        nodes.push(node as SceneNode);
      }
    }
    if (nodes.length > 0) {
      pixso.currentPage.selection = nodes;
    }
  }

  if (msg.type === "go-up") {
    // Select the parent of the current selection (like Shift+Enter)
    const sel = pixso.currentPage.selection;
    if (sel && sel.length > 0) {
      const parent = sel[0].parent;
      if (parent && parent.type !== "PAGE" && parent.type !== "DOCUMENT" && "type" in parent) {
        pixso.currentPage.selection = [parent as SceneNode];
      }
    }
  }

  if (msg.type === "resolve-swap-options") {
    // Lazy resolve: only when user clicks on an INSTANCE_SWAP property
    const { propertyName } = msg;
    pixso.ui.postMessage({ type: "swap-loading", propertyName });

    const data = analyzeSelectionSync();
    // Find preferred values for this property
    let keys: string[] = [];
    for (const inst of data.instances) {
      for (const cp of inst.componentProperties) {
        if (cp.name === propertyName && cp.type === "INSTANCE_SWAP" && cp.preferredValues) {
          keys = cp.preferredValues;
          break;
        }
      }
      if (keys.length > 0) break;
    }

    if (keys.length === 0) {
      pixso.ui.postMessage({ type: "swap-options", swapOptionsMap: {} });
      return;
    }

    Promise.all(
      keys.map(async (key) => {
        try {
          const comp = await pixso.importComponentByKeyAsync(key);
          return { id: comp.id, name: comp.name };
        } catch {
          return null;
        }
      })
    ).then((results) => {
      const options: SwapOption[] = [];
      for (const r of results) {
        if (r) options.push(r);
      }
      const swapOptionsMap: { [name: string]: SwapOption[] } = {};
      if (options.length > 0) {
        swapOptionsMap[propertyName] = options;
      }
      pixso.ui.postMessage({ type: "swap-options", swapOptionsMap });
    });
  }

  if (msg.type === "search-components") {
    // Search across all subscribed libraries for components matching query
    const { query, propertyName, instanceId, componentName: bulkComponentName } = msg;

    pixso.getLibraryListAsync().then(async (libraries) => {
      const results: SwapOption[] = [];
      const q = query.toLowerCase();

      // Also search local components
      const localComponents: SceneNode[] = [];
      function findLocalComponents(node: BaseNode) {
        if ("type" in node) {
          if (node.type === "COMPONENT") {
            localComponents.push(node as SceneNode);
          } else if (node.type === "COMPONENT_SET") {
            // Add individual variants
            if ("children" in node) {
              for (const child of (node as any).children) {
                if (child.type === "COMPONENT") {
                  localComponents.push(child);
                }
              }
            }
          }
          if ("children" in node && node.type !== "INSTANCE") {
            for (const child of (node as any).children) {
              findLocalComponents(child);
            }
          }
        }
      }
      findLocalComponents(pixso.currentPage);

      for (const comp of localComponents) {
        if (comp.name.toLowerCase().includes(q)) {
          results.push({ id: comp.id, name: comp.name });
        }
        if (results.length >= 30) break;
      }

      // Search subscribed libraries
      for (const lib of libraries) {
        if (!lib.subscribed) continue;
        if (results.length >= 30) break;

        try {
          const assets = await pixso.getLibraryByKeyAsync(lib.key);
          for (const comp of assets.componentList) {
            if (results.length >= 30) break;
            if (comp.name.toLowerCase().includes(q)) {
              if (comp.type === "COMPONENT_SET") {
                // Add variants
                for (const v of comp.variants) {
                  if (results.length >= 30) break;
                  results.push({ id: v.key, name: comp.name + " / " + v.name });
                }
              } else {
                results.push({ id: comp.key, name: comp.name });
              }
            }
          }
        } catch {
          // skip unavailable libraries
        }
      }

      pixso.ui.postMessage({
        type: "search-results",
        results,
        propertyName,
        instanceId: instanceId || null,
        componentName: bulkComponentName || null,
      });
    });
  }

  if (msg.type === "apply-swap-from-search") {
    // Apply a swap from search results — may need import if it's a library key
    const { instanceId, propertyName, componentIdOrKey, bulkComponentName } = msg;

    const applySwap = (compId: string) => {
      if (bulkComponentName) {
        // Bulk
        const data = analyzeSelectionSync();
        const targets = data.groupedByComponent[bulkComponentName] ?? [];
        for (const inst of targets) {
          const node = pixso.getNodeById(inst.id) as InstanceNode | null;
          if (node && node.type === "INSTANCE") {
            node.setProperties({ [propertyName]: compId });
          }
        }
      } else if (instanceId) {
        const node = pixso.getNodeById(instanceId) as InstanceNode | null;
        if (node && node.type === "INSTANCE") {
          node.setProperties({ [propertyName]: compId });
        }
      }
      sendSelectionData();
    };

    // Try as local node ID first
    const localNode = pixso.getNodeById(componentIdOrKey);
    if (localNode) {
      applySwap(localNode.id);
    } else {
      // It's a library key — try import (may fail for some libraries)
      try {
        pixso.importComponentByKeyAsync(componentIdOrKey).then((comp) => {
          applySwap(comp.id);
        }).catch(() => {
          console.warn("[Swap] Failed to import component:", componentIdOrKey);
          pixso.notify("Could not import component from library", { error: true });
        });
      } catch {
        console.warn("[Swap] importComponentByKeyAsync threw:", componentIdOrKey);
      }
    }
  }

  if (msg.type === "get-thumbnails") {
    // Export small PNG thumbnails for given node IDs
    const { nodeIds } = msg;
    const results: { id: string; dataUrl: string }[] = [];

    Promise.all(
      (nodeIds as string[]).map(async (id) => {
        try {
          const node = pixso.getNodeById(id);
          if (node && "exportAsync" in node) {
            const bytes = await (node as any).exportAsync({
              format: "PNG",
              constraint: { type: "HEIGHT", value: 32 },
            } as ExportSettingsImage);
            const base64 = pixso.base64Encode(bytes);
            return { id, dataUrl: `data:image/png;base64,${base64}` };
          }
        } catch {
          // skip
        }
        return null;
      })
    ).then((all) => {
      const thumbnails: { [id: string]: string } = {};
      for (const r of all) {
        if (r) thumbnails[r.id] = r.dataUrl;
      }
      pixso.ui.postMessage({ type: "thumbnails", thumbnails });
    });
  }

  if (msg.type === "get-swap-sources") {
    pixso.getLibraryListAsync().then((libraries) => {
      const sources: { key: string; name: string; type: string }[] = [];

      // Check if local components exist
      let hasLocal = false;
      function checkLocal(node: BaseNode) {
        if (hasLocal) return;
        if ("type" in node) {
          if (node.type === "COMPONENT" || node.type === "COMPONENT_SET") {
            hasLocal = true;
            return;
          }
          if ("children" in node && node.type !== "INSTANCE") {
            for (const child of (node as any).children) {
              checkLocal(child);
              if (hasLocal) return;
            }
          }
        }
      }
      checkLocal(pixso.currentPage);

      if (hasLocal) {
        sources.push({ key: "__local__", name: "Local components", type: "local" });
      }

      for (const lib of libraries) {
        if (lib.subscribed) {
          sources.push({ key: lib.key, name: lib.name, type: "library" });
        }
      }

      pixso.ui.postMessage({ type: "swap-sources", sources });
    });
  }

  if (msg.type === "get-library-contents") {
    const { libraryKey } = msg;
    pixso.getLibraryByKeyAsync(libraryKey).then((assets) => {
      const folderMap: { [name: string]: any[] } = {};

      for (const comp of assets.componentList) {
        const folder = comp.containerName || comp.pageName || "Other";
        if (!folderMap[folder]) folderMap[folder] = [];

        if (comp.type === "COMPONENT_SET") {
          folderMap[folder].push({
            key: comp.key,
            name: comp.name,
            thumbnailUrl: comp.thumbnailUrl,
            type: "COMPONENT_SET",
            variants: comp.variants.map((v) => ({
              key: v.key,
              name: v.name,
              thumbnailUrl: v.thumbnailUrl,
            })),
          });
        } else {
          folderMap[folder].push({
            key: comp.key,
            name: comp.name,
            thumbnailUrl: comp.thumbnailUrl,
            type: "COMPONENT",
          });
        }
      }

      const folders = Object.entries(folderMap).map(([name, components]) => ({
        name,
        components,
      }));

      pixso.ui.postMessage({ type: "library-contents", libraryKey, folders });
    });
  }

  if (msg.type === "get-local-components") {
    const folderMap: { [name: string]: { id: string; name: string; type: string; variants?: { id: string; name: string }[] }[] } = {};

    function walkForComponents(node: BaseNode, folderName: string) {
      if ("type" in node) {
        if (node.type === "COMPONENT_SET") {
          const variants = [];
          if ("children" in node) {
            for (const child of (node as any).children) {
              if (child.type === "COMPONENT") {
                variants.push({ id: child.id, name: child.name });
              }
            }
          }
          if (!folderMap[folderName]) folderMap[folderName] = [];
          folderMap[folderName].push({
            id: node.id,
            name: node.name,
            type: "COMPONENT_SET",
            variants,
          });
        } else if (node.type === "COMPONENT") {
          // Only add if not inside a component set
          const parent = node.parent;
          if (!parent || parent.type !== "COMPONENT_SET") {
            if (!folderMap[folderName]) folderMap[folderName] = [];
            folderMap[folderName].push({
              id: node.id,
              name: node.name,
              type: "COMPONENT",
            });
          }
        } else if ("children" in node && node.type !== "INSTANCE") {
          // Use frame name as folder for top-level frames
          const nextFolder = (node.type === "FRAME" || node.type === "SECTION") && node.parent?.type === "PAGE"
            ? node.name
            : folderName;
          for (const child of (node as any).children) {
            walkForComponents(child, nextFolder);
          }
        }
      }
    }

    // Walk ALL pages, not just current
    for (const page of pixso.root.children) {
      walkForComponents(page, page.name);
    }

    const folders = Object.entries(folderMap).map(([name, components]) => ({
      name,
      components,
    }));

    pixso.ui.postMessage({ type: "local-contents", folders });
  }

  if (msg.type === "search-swap-all") {
    const { query } = msg;
    const q = query.toLowerCase();

    pixso.getLibraryListAsync().then(async (libraries) => {
      const results: { id: string; name: string; thumbnailUrl?: string; source: string }[] = [];

      // Search local components
      function searchLocal(node: BaseNode) {
        if (results.length >= 30) return;
        if ("type" in node) {
          if (node.type === "COMPONENT") {
            const parent = node.parent;
            if (parent && parent.type === "COMPONENT_SET") {
              if ((parent.name + " / " + node.name).toLowerCase().includes(q)) {
                results.push({ id: node.id, name: parent.name + " / " + node.name, source: "Local" });
              }
            } else if (node.name.toLowerCase().includes(q)) {
              results.push({ id: node.id, name: node.name, source: "Local" });
            }
          } else if (node.type === "COMPONENT_SET") {
            // Search variants inside
            if ("children" in node) {
              for (const child of (node as any).children) {
                if (results.length >= 30) break;
                if (child.type === "COMPONENT" && (node.name + " / " + child.name).toLowerCase().includes(q)) {
                  results.push({ id: child.id, name: node.name + " / " + child.name, source: "Local" });
                }
              }
            }
          }
          if ("children" in node && node.type !== "INSTANCE" && node.type !== "COMPONENT_SET") {
            for (const child of (node as any).children) {
              searchLocal(child);
            }
          }
        }
      }
      searchLocal(pixso.currentPage);

      // Search subscribed libraries
      for (const lib of libraries) {
        if (!lib.subscribed || results.length >= 30) continue;
        try {
          const assets = await pixso.getLibraryByKeyAsync(lib.key);
          for (const comp of assets.componentList) {
            if (results.length >= 30) break;
            if (comp.type === "COMPONENT_SET") {
              if (comp.name.toLowerCase().includes(q)) {
                for (const v of comp.variants) {
                  if (results.length >= 30) break;
                  results.push({
                    id: v.key,
                    name: comp.name + " / " + v.name,
                    thumbnailUrl: v.thumbnailUrl,
                    source: lib.name,
                  });
                }
              }
            } else if (comp.name.toLowerCase().includes(q)) {
              results.push({
                id: comp.key,
                name: comp.name,
                thumbnailUrl: comp.thumbnailUrl,
                source: lib.name,
              });
            }
          }
        } catch {}
      }

      pixso.ui.postMessage({ type: "swap-search-results", results });
    });
  }

  if (msg.type === "get-preferred-swap-values") {
    const { propertyName } = msg;

    // Helper: find the current swap component via componentPropertyReferences (recursive)
    function findCurrentSwapComp(): ComponentNode | null {
      const sel = pixso.currentPage.selection;
      if (!sel || sel.length === 0) return null;

      function searchInNode(node: SceneNode): ComponentNode | null {
        if ("componentPropertyReferences" in node) {
          const refs = (node as any).componentPropertyReferences;
          if (refs && refs.mainComponent === propertyName && isInstanceNode(node)) {
            return node.mainComponent;
          }
        }
        if (hasChildren(node)) {
          for (const child of node.children) {
            const found = searchInNode(child);
            if (found) return found;
          }
        }
        return null;
      }

      for (const node of sel) {
        const found = searchInNode(node);
        if (found) return found;
      }
      return null;
    }



    // Step 1: Get preferredValues keys
    let prefKeys: string[] = [];
    const sel = pixso.currentPage.selection;
    if (sel && sel.length > 0) {
      const stack: SceneNode[] = [...sel];
      while (stack.length > 0 && prefKeys.length === 0) {
        const node = stack.pop()!;
        if (isInstanceNode(node)) {
          const mc = node.mainComponent;
          if (mc) {
            let defs: ComponentPropertyDefinitions | null = null;
            const p = mc.parent;
            if (p && p.type === "COMPONENT_SET") {
              defs = (p as ComponentSetNode).componentPropertyDefinitions;
            } else {
              defs = mc.componentPropertyDefinitions;
            }
            if (defs && defs[propertyName] && defs[propertyName].preferredValues) {
              prefKeys = (defs[propertyName].preferredValues as any[]).map((pv: any) => pv.key);
            }
          }
        }
        if (prefKeys.length === 0 && hasChildren(node)) {
          for (const c of node.children) stack.push(c);
        }
      }
    }

    // Step 2: Find current swap component
    let currentComp = findCurrentSwapComp();

    // Fallback: try componentProperties value via getNodeById
    if (!currentComp && sel && sel.length > 0) {
      for (const node of sel) {
        if (isInstanceNode(node)) {
          const cp = node.componentProperties;
          if (cp && cp[propertyName] && cp[propertyName].type === "INSTANCE_SWAP") {
            const val = cp[propertyName].value as string;
            const valNode = pixso.getNodeById(val);
            if (valNode) {
              if (valNode.type === "COMPONENT") {
                currentComp = valNode as ComponentNode;
              } else if (valNode.type === "INSTANCE") {
                currentComp = (valNode as InstanceNode).mainComponent;
              }
            }
            break;
          }
        }
      }
    }

    console.log("[Swap] property:", propertyName, "prefKeys:", prefKeys.length, "currentComp:", currentComp?.name || "null");

    let currentKeyInPreferred = false;
    if (currentComp && prefKeys.length > 0) {
      currentKeyInPreferred = prefKeys.includes(currentComp.key);
    }


    // Decision: show quick swap OR navigate to current component's location
    const shouldShowQuickSwap = prefKeys.length > 0 && currentKeyInPreferred;

    if (shouldShowQuickSwap) {
      // Show quick swap — use currentComp's library only (fast, targeted)
      const values: { id: string; name: string; thumbnailDataUrl: string }[] = [];
      const keySet = new Set(prefKeys);

      // Get library info from currentComp directly
      currentComp!.getLibraryInfoAsync().then(async (libInfo) => {
        if (libInfo && libInfo.key) {
          try {
            const assets = await pixso.getLibraryByKeyAsync(libInfo.key);
            for (const comp of assets.componentList) {
              if (comp.type === "COMPONENT_SET") {
                if (keySet.has(comp.key)) {
                  const first = comp.variants[0];
                  values.push({ id: first?.key || comp.key, name: comp.name, thumbnailDataUrl: comp.thumbnailUrl || "" });
                  keySet.delete(comp.key);
                }
                for (const v of comp.variants) {
                  if (keySet.has(v.key)) {
                    values.push({ id: v.key, name: comp.name + " / " + v.name, thumbnailDataUrl: v.thumbnailUrl || "" });
                    keySet.delete(v.key);
                  }
                }
              } else if (keySet.has(comp.key)) {
                values.push({ id: comp.key, name: comp.name, thumbnailDataUrl: comp.thumbnailUrl || "" });
                keySet.delete(comp.key);
              }
            }
          } catch {}
        }

        // For remaining keys, try other subscribed libraries
        if (keySet.size > 0) {
          try {
            const libraries = await pixso.getLibraryListAsync();
            for (const lib of libraries) {
              if (!lib.subscribed || keySet.size === 0) continue;
              if (libInfo && lib.key === libInfo.key) continue; // already searched
              try {
                const assets = await pixso.getLibraryByKeyAsync(lib.key);
                for (const comp of assets.componentList) {
                  if (comp.type === "COMPONENT_SET") {
                    for (const v of comp.variants) {
                      if (keySet.has(v.key)) {
                        values.push({ id: v.key, name: comp.name + " / " + v.name, thumbnailDataUrl: v.thumbnailUrl || "" });
                        keySet.delete(v.key);
                      }
                    }
                  } else if (keySet.has(comp.key)) {
                    values.push({ id: comp.key, name: comp.name, thumbnailDataUrl: comp.thumbnailUrl || "" });
                    keySet.delete(comp.key);
                  }
                  if (keySet.size === 0) break;
                }
              } catch {}
            }
          } catch {}
        }

        console.log("[Swap] Quick swap resolved:", values.length, "of", prefKeys.length);
        pixso.ui.postMessage({ type: "preferred-swap-values", propertyName, values });
      }).catch(() => {
        // getLibraryInfoAsync failed — send empty
        console.log("[Swap] getLibraryInfoAsync failed, sending empty");
        pixso.ui.postMessage({ type: "preferred-swap-values", propertyName, values: [] });
      });
    } else {
      // Navigate to current component's location
      if (currentComp) {
        const mcP = currentComp.parent;
        const compName = (mcP && mcP.type === "COMPONENT_SET") ? mcP.name : currentComp.name;
        // All info directly from ComponentNode — no async needed
        const nav: any = {
          componentName: compName,
          componentKey: currentComp.key,
          containerName: (currentComp as any).containerName || "",
          pageName: (currentComp as any).pageName || "",
          remote: currentComp.remote,
        };
        // Also try to get library key (async but fast)
        currentComp.getLibraryInfoAsync().then((libInfo) => {
          if (libInfo && libInfo.key) nav.libraryKey = libInfo.key;
          if (libInfo && libInfo.name) nav.libraryName = libInfo.name;
        }).catch(() => {}).finally(() => {
          console.log("[Swap] navigateTo:", JSON.stringify(nav));
          pixso.ui.postMessage({
            type: "preferred-swap-values",
            propertyName,
            values: [],
            navigateTo: nav,
          });
        });
        return;
      } else {
        pixso.ui.postMessage({ type: "preferred-swap-values", propertyName, values: [] });
      }
    }
  }

  if (msg.type === "refresh") {
    sendSelectionData();
  }
});

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

  for (const [propName, propValue] of Object.entries(compProps)) {
    const def = definitions?.[propName];
    const info: ComponentPropertyInfo = {
      name: propName,
      type: propValue.type,
      currentValue: propValue.value,
      defaultValue: def?.defaultValue,
    };

    if (propValue.type === "VARIANT" && def) {
      info.options = (def as any).variantOptions ?? [];
    }

    if (propValue.type === "INSTANCE_SWAP" && typeof propValue.value === "string") {
      // Resolve component ID to name
      const swapNode = pixso.getNodeById(propValue.value);
      if (swapNode) {
        info.currentValueName = swapNode.name;
      } else {
        info.currentValueName = String(propValue.value);
      }

      // Collect preferred values keys for async resolution
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

// Send data immediately (used after property changes — sync, no loading)
function sendSelectionData() {
  const data = analyzeSelectionSync();
  const parentInfo = getParentInfo();
  pixso.ui.postMessage({
    type: "selection-data",
    data,
    parentInfo,
  });
  resolveSwapOptions(data);
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
    resolveSwapOptions(data);
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

  if (msg.type === "refresh") {
    sendSelectionData();
  }
});

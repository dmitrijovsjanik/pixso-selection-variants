/// <reference path="../node_modules/@pixso/plugin-typings/index.d.ts" />

// Selection Variants & Instances — Pixso Plugin (Sandbox)
// This runs in the plugin sandbox with access to the Pixso API.

// ─── Types ───────────────────────────────────────────────

interface VariantProperty {
  name: string;
  currentValue: string;
  options: string[];
}

interface ComponentPropertyInfo {
  name: string;
  type: string; // "BOOLEAN" | "TEXT" | "INSTANCE_SWAP" | "VARIANT"
  currentValue: string | boolean;
  defaultValue?: string | boolean;
  preferredValues?: string[];
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

    if (def && "preferredValues" in def) {
      info.preferredValues = (def as any).preferredValues;
    }

    props.push(info);
  }

  return props;
}

function collectInstances(
  nodes: readonly SceneNode[],
  depth: number = 0,
  results: InstanceInfo[] = []
): InstanceInfo[] {
  for (const node of nodes) {
    if (isInstanceNode(node)) {
      const mainComponent = node.mainComponent;
      const componentName = mainComponent?.name ?? "Unknown Component";
      const componentId = mainComponent?.id ?? null;

      results.push({
        id: node.id,
        name: node.name,
        componentName,
        componentId,
        depth,
        variantProperties: getVariantProperties(node),
        componentProperties: getComponentProperties(node),
        path: getNodePath(node),
      });

      // Also recurse into nested instances
      if (hasChildren(node)) {
        collectInstances(node.children, depth + 1, results);
      }
    } else if (hasChildren(node)) {
      collectInstances(node.children, depth + 1, results);
    }
  }
  return results;
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

function analyzeSelection(): SelectionData {
  const selection = pixso.currentPage.selection;
  if (!selection || selection.length === 0) {
    return {
      instances: [],
      groupedByComponent: {},
      totalSelected: 0,
      hasSelection: false,
    };
  }

  // Collect all instances including top-level selected instances
  const allInstances: InstanceInfo[] = [];

  for (const node of selection) {
    if (isInstanceNode(node)) {
      // The selected node itself is an instance — include it at depth 0
      const mainComponent = node.mainComponent;
      allInstances.push({
        id: node.id,
        name: node.name,
        componentName: mainComponent?.name ?? "Unknown Component",
        componentId: mainComponent?.id ?? null,
        depth: 0,
        variantProperties: getVariantProperties(node),
        componentProperties: getComponentProperties(node),
        path: getNodePath(node),
      });
      // Also recurse into children
      if (hasChildren(node)) {
        collectInstances(node.children, 1, allInstances);
      }
    } else if (hasChildren(node)) {
      collectInstances(node.children, 0, allInstances);
    }
  }

  return {
    instances: allInstances,
    groupedByComponent: groupByComponent(allInstances),
    totalSelected: selection.length,
    hasSelection: true,
  };
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

// Selection history for back navigation
let previousSelection: string[] = [];
let ignoreNextSelectionChange = false;

function savePreviousSelection() {
  const sel = pixso.currentPage.selection;
  if (sel && sel.length > 0) {
    previousSelection = sel.map((n) => n.id);
  }
}

// Send initial data
function sendSelectionData() {
  const data = analyzeSelection();
  const hasPrevious = previousSelection.length > 0;
  pixso.ui.postMessage({ type: "selection-data", data, hasPreviousSelection: hasPrevious });
}

sendSelectionData();

// Listen for selection changes
pixso.on("selectionchange", () => {
  if (ignoreNextSelectionChange) {
    ignoreNextSelectionChange = false;
    sendSelectionData();
    return;
  }
  // User changed selection manually — clear history
  previousSelection = [];
  sendSelectionData();
});

// Listen for messages from UI
pixso.ui.on("message", (msg: any) => {
  if (msg.type === "set-variant") {
    // Switch a variant property on a specific instance
    const { instanceId, propertyName, newValue } = msg;
    const instanceNode = pixso.getNodeById(instanceId) as InstanceNode | null;
    if (!instanceNode || instanceNode.type !== "INSTANCE") return;

    // Find the instance info to get current variant props
    const data = analyzeSelection();
    const instInfo = data.instances.find((i) => i.id === instanceId);
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
    const data = analyzeSelection();
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
    const data = analyzeSelection();
    const targetInstances = data.groupedByComponent[componentName] ?? [];

    for (const instInfo of targetInstances) {
      const instanceNode = pixso.getNodeById(instInfo.id) as InstanceNode | null;
      if (!instanceNode || instanceNode.type !== "INSTANCE") continue;

      instanceNode.setProperties({ [propertyName]: newValue });
    }

    sendSelectionData();
  }

  if (msg.type === "select-instance") {
    // Select a specific instance in the canvas
    const { instanceId } = msg;
    const node = pixso.getNodeById(instanceId);
    if (node && "type" in node) {
      savePreviousSelection();
      ignoreNextSelectionChange = true;
      pixso.currentPage.selection = [node as SceneNode];
    }
  }

  if (msg.type === "select-instances-by-component") {
    // Select all instances of a specific component
    const { componentName } = msg;
    const data = analyzeSelection();
    const targetInstances = data.groupedByComponent[componentName] ?? [];
    const nodes: SceneNode[] = [];
    for (const inst of targetInstances) {
      const node = pixso.getNodeById(inst.id);
      if (node && "type" in node) {
        nodes.push(node as SceneNode);
      }
    }
    if (nodes.length > 0) {
      savePreviousSelection();
      ignoreNextSelectionChange = true;
      pixso.currentPage.selection = nodes;
    }
  }

  if (msg.type === "restore-selection") {
    // Go back to previous selection
    if (previousSelection.length > 0) {
      const nodes: SceneNode[] = [];
      for (const id of previousSelection) {
        const node = pixso.getNodeById(id);
        if (node && "type" in node) {
          nodes.push(node as SceneNode);
        }
      }
      previousSelection = [];
      if (nodes.length > 0) {
        ignoreNextSelectionChange = true;
        pixso.currentPage.selection = nodes;
      }
    }
  }

  if (msg.type === "refresh") {
    sendSelectionData();
  }
});

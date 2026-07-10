import {
  box,
  column,
  grid,
  layers,
  node,
  paragraph,
  row,
  rule,
  run,
  shape,
} from "../index.mjs";

export const Fragment = "fragment";

function normalizeProps(props = {}) {
  if (!props) return { props: {}, children: [] };
  const { children, key, ref, ...rest } = props;
  return {
    props: key == null ? rest : { ...rest, key },
    children: children == null ? [] : Array.isArray(children) ? children : [children],
  };
}

export function createElement(type, props = {}) {
  const normalized = normalizeProps(props);
  if (typeof type === "function") {
    return type({ ...normalized.props, children: normalized.children });
  }
  if (type === Fragment) {
    return node("fragment", normalized.props, normalized.children);
  }
  return node(type, normalized.props, normalized.children);
}

export {
  box,
  column,
  grid,
  layers,
  node,
  paragraph,
  row,
  rule,
  run,
  shape,
};

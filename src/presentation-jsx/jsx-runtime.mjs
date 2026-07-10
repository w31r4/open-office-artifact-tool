import { createElement, Fragment } from "./index.mjs";

export { Fragment };

export function jsx(type, props, key) {
  return createElement(type, key == null ? props : { ...props, key });
}

export function jsxs(type, props, key) {
  return createElement(type, key == null ? props : { ...props, key });
}

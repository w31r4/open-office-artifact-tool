import { createElement, Fragment } from "./index.mjs";

export { Fragment };

export function jsxDEV(type, props, key) {
  return createElement(type, key == null ? props : { ...props, key });
}

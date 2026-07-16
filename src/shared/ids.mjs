let nextId = 1;

export function aid(prefix) {
  return `${prefix}/${(nextId++).toString(36).padStart(4, "0")}`;
}

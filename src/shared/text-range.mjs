export function createTextRange(parent, id, options = {}) {
  const getText = options.getText || (() => parent.text?.value ?? parent.text ?? parent.display ?? "");
  const setText = options.setText || ((value) => {
    if (parent.text && typeof parent.text.set === "function") parent.text.set(value);
    else if (parent.text && typeof parent.text === "object" && "value" in parent.text) parent.text.value = String(value ?? "");
    else if ("text" in parent) parent.text = String(value ?? "");
    else if ("display" in parent) parent.display = String(value ?? "");
  });
  const replaceText = options.replace || ((search, replacement) => {
    const next = String(getText()).replace(search, replacement);
    setText(next);
  });
  return {
    kind: "textRange",
    id,
    parentId: parent.id,
    parentKind: options.parentKind || parent.kind || parent.constructor?.name,
    get text() { return getText(); },
    set text(value) { setText(value); },
    replace(search, replacement) { replaceText(search, replacement); return this; },
  };
}

export function textRangeRecord(parent, options = {}) {
  const range = createTextRange(parent, `${parent.id}/text`, options);
  const text = String(range.text || "");
  return { kind: "textRange", id: range.id, parentId: parent.id, parentKind: range.parentKind, text, textPreview: text.slice(0, 300), textChars: text.length, ...(options.record || {}) };
}

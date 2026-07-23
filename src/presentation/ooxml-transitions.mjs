const TRANSITION_EFFECTS = new Set(["fade", "push"]);
const TRANSITION_SPEEDS = new Set(["slow", "medium", "fast"]);
const TRANSITION_DIRECTIONS = new Set(["left", "up", "right", "down"]);
const TRANSITION_KEYS = new Set(["effect", "direction", "speed", "advanceOnClick", "advanceAfterMs"]);
const MAX_ADVANCE_AFTER_MS = 86_400_000;

export const PRESENTATION_TRANSITION_CAPABILITY = Symbol.for("open-office-artifact-tool.open-chestnut-slide-transition-capability");

function own(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function normalizeEffect(value) {
  const effect = String(value || "").trim().toLowerCase();
  if (!TRANSITION_EFFECTS.has(effect)) {
    throw new TypeError("Presentation transition effect must be fade or push.");
  }
  return effect;
}

function normalizeSpeed(value) {
  const speed = String(value ?? "medium").trim().toLowerCase();
  if (!TRANSITION_SPEEDS.has(speed)) {
    throw new TypeError("Presentation transition speed must be slow, medium, or fast.");
  }
  return speed;
}

function normalizeAdvanceAfter(value) {
  const milliseconds = Number(value);
  if (!Number.isSafeInteger(milliseconds) || milliseconds < 0 || milliseconds > MAX_ADVANCE_AFTER_MS) {
    throw new RangeError(`Presentation transition advanceAfterMs must be an integer from 0 through ${MAX_ADVANCE_AFTER_MS}.`);
  }
  return milliseconds;
}

export function normalizePresentationTransition(config) {
  if (!config || typeof config !== "object" || Array.isArray(config)) {
    throw new TypeError("Presentation transition must be an object.");
  }
  const unsupported = Object.keys(config).filter((key) => !TRANSITION_KEYS.has(key));
  if (unsupported.length) {
    throw new TypeError(`Presentation transition has unsupported fields: ${unsupported.join(", ")}.`);
  }
  const effect = normalizeEffect(config.effect);
  const speed = normalizeSpeed(config.speed);
  const transition = { effect, speed };
  if (effect === "push") {
    const direction = String(config.direction ?? "left").trim().toLowerCase();
    if (!TRANSITION_DIRECTIONS.has(direction)) {
      throw new TypeError("Presentation push transition direction must be left, up, right, or down.");
    }
    transition.direction = direction;
  } else if (own(config, "direction") && config.direction != null) {
    throw new TypeError("Presentation fade transition does not accept direction.");
  }
  if (own(config, "advanceOnClick") && typeof config.advanceOnClick !== "boolean") {
    throw new TypeError("Presentation transition advanceOnClick must be a boolean.");
  }
  transition.advanceOnClick = config.advanceOnClick ?? true;
  if (own(config, "advanceAfterMs") && config.advanceAfterMs != null) {
    transition.advanceAfterMs = normalizeAdvanceAfter(config.advanceAfterMs);
  }
  return transition;
}

function cloneTransition(value) {
  return value ? { ...value } : undefined;
}

export class SlideTransition {
  constructor(slide, config) {
    this.slide = slide;
    this._value = config == null ? undefined : normalizePresentationTransition(config);
  }

  get id() { return `${this.slide.id}/transition`; }
  get configured() { return Boolean(this._value); }
  get effect() { return this._value?.effect; }
  get direction() { return this._value?.direction; }
  get speed() { return this._value?.speed; }
  get advanceOnClick() { return this._value?.advanceOnClick; }
  get advanceAfterMs() { return this._value?.advanceAfterMs; }
  get capability() {
    const imported = this[PRESENTATION_TRANSITION_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, partPresent: this.configured, editable: true, addable: true };
  }

  set(config) {
    const capability = this.capability;
    if (capability.sourceBound && !capability.editable) {
      throw new Error("Presentation slide transition is source-bound and cannot be semantically replaced by this codec profile.");
    }
    this._value = normalizePresentationTransition(config);
    return this;
  }

  clear() {
    const capability = this.capability;
    if (capability.sourceBound && capability.partPresent && !capability.editable) {
      throw new Error("Presentation slide transition is source-bound and cannot be removed by this codec profile.");
    }
    this._value = undefined;
    return this;
  }

  inspectRecord() {
    return {
      kind: "transition",
      id: this.id,
      slide: this.slide.index + 1,
      configured: this.configured,
      ...(this._value || {}),
      capability: this.capability,
    };
  }

  toJSON() { return cloneTransition(this._value); }

  // The OpenChestnut adapter uses this after it has decoded a validated
  // protobuf payload. It deliberately bypasses public source-bound mutation
  // checks; callers use set()/clear(), which remain capability-aware.
  _setImported(config) {
    this._value = config == null ? undefined : normalizePresentationTransition(config);
    return this;
  }
}

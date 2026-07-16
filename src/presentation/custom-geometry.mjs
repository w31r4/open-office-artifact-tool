const MAX_PATHS = 64;
const MAX_COMMANDS = 16_384;
const MAX_COORDINATE = 2_147_483_647;

function coordinate(value, label) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < -MAX_COORDINATE || number > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a safe integer within the DrawingML signed 32-bit coordinate range.`);
  }
  return number;
}

function point(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unknown = Object.keys(value).filter((key) => key !== "x" && key !== "y");
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  return { x: coordinate(value.x, `${label}.x`), y: coordinate(value.y, `${label}.y`) };
}

function command(value, pathIndex, commandIndex) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must be an object.`);
  const keys = Object.keys(value);
  if (keys.length !== 1) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must contain exactly one command.`);
  const label = `Presentation custom path ${pathIndex + 1} command ${commandIndex + 1}`;
  if (keys[0] === "moveTo" || keys[0] === "lineTo") return { [keys[0]]: point(value[keys[0]], `${label}.${keys[0]}`) };
  if (keys[0] === "cubicBezTo") {
    const cubic = value.cubicBezTo;
    if (!cubic || typeof cubic !== "object" || Array.isArray(cubic)) throw new TypeError(`${label}.cubicBezTo must be an object.`);
    const unknown = Object.keys(cubic).filter((key) => !new Set(["x1", "y1", "x2", "y2", "x", "y"]).has(key));
    if (unknown.length) throw new TypeError(`${label}.cubicBezTo has unsupported fields: ${unknown.join(", ")}.`);
    return {
      cubicBezTo: {
        x1: coordinate(cubic.x1, `${label}.cubicBezTo.x1`),
        y1: coordinate(cubic.y1, `${label}.cubicBezTo.y1`),
        x2: coordinate(cubic.x2, `${label}.cubicBezTo.x2`),
        y2: coordinate(cubic.y2, `${label}.cubicBezTo.y2`),
        x: coordinate(cubic.x, `${label}.cubicBezTo.x`),
        y: coordinate(cubic.y, `${label}.cubicBezTo.y`),
      },
    };
  }
  if (keys[0] === "close") {
    if (value.close !== true && (typeof value.close !== "object" || value.close == null || Array.isArray(value.close) || Object.keys(value.close).length)) {
      throw new TypeError(`${label}.close must be true or an empty object.`);
    }
    return { close: {} };
  }
  throw new TypeError(`${label} uses unsupported command ${keys[0]}.`);
}

export function normalizePresentationCustomPaths(value, { geometry } = {}) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_PATHS) throw new RangeError(`Presentation custom geometry must contain 1 through ${MAX_PATHS} paths.`);
  let commandCount = 0;
  return value.map((path, pathIndex) => {
    if (!path || typeof path !== "object" || Array.isArray(path)) throw new TypeError(`Presentation custom path ${pathIndex + 1} must be an object.`);
    const unknown = Object.keys(path).filter((key) => !new Set(["width", "height", "commands"]).has(key));
    if (unknown.length) throw new TypeError(`Presentation custom path ${pathIndex + 1} has unsupported fields: ${unknown.join(", ")}.`);
    const width = coordinate(path.width, `Presentation custom path ${pathIndex + 1}.width`);
    const height = coordinate(path.height, `Presentation custom path ${pathIndex + 1}.height`);
    if (width <= 0 || height <= 0) throw new RangeError(`Presentation custom path ${pathIndex + 1} width and height must be positive.`);
    if (!Array.isArray(path.commands) || path.commands.length === 0) throw new TypeError(`Presentation custom path ${pathIndex + 1} requires commands.`);
    commandCount += path.commands.length;
    if (commandCount > MAX_COMMANDS) throw new RangeError(`Presentation custom geometry exceeds the ${MAX_COMMANDS}-command budget.`);
    return { width, height, commands: path.commands.map((item, commandIndex) => command(item, pathIndex, commandIndex)) };
  });
}

export function presentationCustomPathsSvg(paths, frame, { escape = String } = {}) {
  return paths.map((path) => {
    const commands = path.commands.map((item) => {
      if (item.moveTo) return `M ${item.moveTo.x} ${item.moveTo.y}`;
      if (item.lineTo) return `L ${item.lineTo.x} ${item.lineTo.y}`;
      if (item.cubicBezTo) return `C ${item.cubicBezTo.x1} ${item.cubicBezTo.y1} ${item.cubicBezTo.x2} ${item.cubicBezTo.y2} ${item.cubicBezTo.x} ${item.cubicBezTo.y}`;
      return "Z";
    }).join(" ");
    return `<path d="${escape(commands)}" transform="translate(${frame.left} ${frame.top}) scale(${frame.width / path.width} ${frame.height / path.height})"/>`;
  }).join("");
}

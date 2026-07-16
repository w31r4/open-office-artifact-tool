import { analyzePdfReadingOrder } from "./reading-order.mjs";

function logicalTableGroups(pages = []) {
  const groups = new Map();
  pages.forEach((page, pageIndex) => page.tables.forEach((table) => {
    if (!table.semanticId) return;
    if (!groups.has(table.semanticId)) groups.set(table.semanticId, []);
    groups.get(table.semanticId).push({ page, pageIndex, table });
  }));
  return groups;
}

export function pdfCrossPageTableIssues(pages = []) {
  const issues = [];
  for (const [semanticId, segments] of logicalTableGroups(pages)) {
    if (segments.length < 2) continue;
    const columns = segments.map(({ table }) => table.grid().columns);
    if (new Set(columns).size !== 1) {
      issues.push({ code: "crossPageTableColumnMismatch", semanticId, message: `Logical PDF table ${semanticId} has incompatible segment column counts ${columns.join(", ")}.` });
    }
    const pageIndexes = segments.map(({ pageIndex }) => pageIndex);
    if (pageIndexes.some((pageIndex, index) => index > 0 && pageIndex !== pageIndexes[index - 1] + 1)) {
      issues.push({ code: "crossPageTablePageGap", semanticId, message: `Logical PDF table ${semanticId} must use consecutive pages.` });
    }
    segments.forEach(({ page, pageIndex, table }, segmentIndex) => {
      const order = analyzePdfReadingOrder(page).declaredIds;
      const position = order.indexOf(table.id);
      if (position < 0) return;
      if (segmentIndex > 0 && position !== 0) {
        issues.push({ code: "crossPageTableInterleaving", semanticId, page: pageIndex + 1, id: table.id, message: `Continuation segment ${table.id} must be the first semantic item on page ${pageIndex + 1}.` });
      }
      if (segmentIndex < segments.length - 1 && position !== order.length - 1) {
        issues.push({ code: "crossPageTableInterleaving", semanticId, page: pageIndex + 1, id: table.id, message: `Non-final segment ${table.id} must be the last semantic item on page ${pageIndex + 1}.` });
      }
    });
  }
  return issues;
}

export function mergePdfSemanticTableGroups(plans = []) {
  const occurrences = new Map();
  plans.forEach((plan) => plan.semanticGroups.forEach((group) => {
    if (group.role !== "Table" || !group.semanticId) return;
    if (!occurrences.has(group.semanticId)) occurrences.set(group.semanticId, []);
    occurrences.get(group.semanticId).push({ plan, group });
  }));
  for (const [semanticId, segments] of occurrences) {
    if (segments.length < 2) continue;
    const [first, ...continuations] = segments;
    const merged = {
      role: "Table",
      title: first.group.title,
      sourceId: first.group.sourceId,
      structureId: semanticId,
      semanticId,
      spansPages: true,
      children: segments.flatMap(({ group }) => group.children || []),
    };
    first.plan.semanticGroups = first.plan.semanticGroups.map((group) => group === first.group ? merged : group);
    for (const { plan, group } of continuations) plan.semanticGroups = plan.semanticGroups.filter((candidate) => candidate !== group);
  }
  return plans;
}

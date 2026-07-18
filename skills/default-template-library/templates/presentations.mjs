import { Presentation } from "open-office-artifact-tool";

import { addPresentationText, noLine } from "./common.mjs";

const titleStyle = Object.freeze({ fontFamily: "Aptos Display", fontSize: 34, bold: true, color: "#0F3D4C" });

function createPresentation() {
  return Presentation.create({ slideSize: { width: 1280, height: 720 } });
}

function addAccent(slide, fill = "#0F766E") {
  slide.shapes.add({
    name: "title-accent",
    geometry: "rect",
    position: { left: 0, top: 0, width: 1280, height: 20 },
    fill,
    line: noLine(),
  });
}

function addCard(slide, { name, left, top, width, height, heading, body, fill = "#E8F5F3", line = "#91C7C0", color = "#0F3D4C" }) {
  slide.shapes.add({
    name,
    geometry: "roundRect",
    position: { left, top, width, height },
    fill,
    line: { fill: line, width: 1 },
    text: `${heading}\n\n${body}`,
    textStyle: { fontFamily: "Aptos", fontSize: 18, bold: true, color },
  });
}

export function buildProjectKickoff() {
  const presentation = createPresentation();
  const title = presentation.slides.add({
    name: "Kickoff overview",
    notes: "Introduce the outcome first, then the commitment requested from the group.",
    background: { fill: "#F4FBFA" },
  });
  addAccent(title);
  addPresentationText(title, {
    name: "kickoff-title",
    text: "Project Kickoff",
    position: { left: 84, top: 108, width: 860, height: 92 },
    textStyle: { fontFamily: "Aptos Display", fontSize: 48, bold: true, color: "#0F3D4C" },
  });
  addPresentationText(title, {
    name: "kickoff-subtitle",
    text: "A source-free working deck for outcome, scope, ownership, and decisions",
    position: { left: 88, top: 220, width: 760, height: 76 },
    textStyle: { fontFamily: "Aptos", fontSize: 21, color: "#48656A" },
  });
  title.shapes.add({
    name: "outcome-panel",
    geometry: "roundRect",
    position: { left: 86, top: 346, width: 1090, height: 196 },
    fill: "#D9F0EE",
    line: { fill: "#88B9B3", width: 1 },
    text: "OUTCOME\nDeliver a measurable first release while preserving a clear rollback path and a shared record of decisions.",
    textStyle: { fontFamily: "Aptos", fontSize: 22, bold: true, color: "#173B3D" },
  });

  const scope = presentation.slides.add({
    name: "Scope and plan",
    notes: "Confirm the boundary before discussing detailed implementation tasks.",
    background: { fill: "#FFFFFF" },
  });
  addPresentationText(scope, {
    name: "scope-title",
    text: "Scope and plan",
    position: { left: 80, top: 64, width: 700, height: 60 },
    textStyle: titleStyle,
  });
  const planCards = [
    ["In scope", "Outcome, decision record, bounded delivery, and verification evidence."],
    ["Not in scope", "Unbounded redesigns, unspecified integrations, and unowned operational work."],
    ["First checkpoint", "Review evidence, risks, and the next irreversible decision."],
  ];
  planCards.forEach(([heading, body], index) => addCard(scope, {
    name: `scope-card-${index + 1}`,
    left: 80 + index * 385,
    top: 210,
    width: 330,
    height: 282,
    heading,
    body,
    fill: index === 1 ? "#FEF3C7" : "#E8F5F3",
    line: index === 1 ? "#E8B854" : "#91C7C0",
  }));

  const operating = presentation.slides.add({
    name: "Owners and decisions",
    notes: "Close by making ownership and the decision cadence explicit.",
    background: { fill: "#0F3D4C" },
  });
  addPresentationText(operating, {
    name: "operating-title",
    text: "Owners and decisions",
    position: { left: 80, top: 66, width: 790, height: 60 },
    textStyle: { ...titleStyle, color: "#FFFFFF" },
  });
  const rows = [
    ["Outcome owner", "Owns priorities, success criteria, and the decision record."],
    ["Delivery owner", "Owns the plan, risk surfacing, and evidence collection."],
    ["Decision cadence", "Review weekly; escalate only a named blocking decision."],
  ];
  rows.forEach(([heading, body], index) => operating.shapes.add({
    name: `operating-row-${index + 1}`,
    geometry: "roundRect",
    position: { left: 80, top: 194 + index * 140, width: 1110, height: 102 },
    fill: "#174F5D",
    line: { fill: "#4D8991", width: 1 },
    text: `${heading}: ${body}`,
    textStyle: { fontFamily: "Aptos", fontSize: 17, bold: true, color: "#FFFFFF" },
  }));
  return presentation;
}

export function buildOperatingReview() {
  const presentation = createPresentation();
  const scorecard = presentation.slides.add({
    name: "Operating scorecard",
    notes: "Start with the observed operating result and state the decision required from the group.",
    background: { fill: "#F8FAFC" },
  });
  addAccent(scorecard, "#1D4ED8");
  addPresentationText(scorecard, {
    name: "operating-review-title",
    text: "Operating Review",
    position: { left: 80, top: 80, width: 760, height: 76 },
    textStyle: { fontFamily: "Aptos Display", fontSize: 46, bold: true, color: "#0F172A" },
  });
  addPresentationText(scorecard, {
    name: "operating-review-subtitle",
    text: "Review results, risks, decisions, and accountable follow-through.",
    position: { left: 84, top: 172, width: 870, height: 76 },
    textStyle: { fontFamily: "Aptos", fontSize: 20, color: "#475569" },
  });
  [
    ["Outcome", "On track", "The named user result is trending within the agreed threshold."],
    ["Delivery", "Watch", "Two dependencies need an owner-confirmed date this week."],
    ["Decision", "Required", "Confirm scope boundary and the one escalation path."],
  ].forEach(([heading, status, body], index) => addCard(scorecard, {
    name: `scorecard-${index + 1}`,
    left: 80 + index * 380,
    top: 300,
    width: 326,
    height: 236,
    heading: `${heading}: ${status}`,
    body,
    fill: index === 1 ? "#FEF3C7" : index === 2 ? "#DBEAFE" : "#DCFCE7",
    line: index === 1 ? "#E8B854" : index === 2 ? "#93C5FD" : "#86C99A",
  }));

  const delivery = presentation.slides.add({
    name: "Delivery and risks",
    notes: "Use this page to distinguish a manageable watch item from a decision-blocking risk.",
    background: { fill: "#FFFFFF" },
  });
  addPresentationText(delivery, {
    name: "delivery-title",
    text: "Delivery and risks",
    position: { left: 80, top: 64, width: 820, height: 60 },
    textStyle: titleStyle,
  });
  [
    ["On track", "Outcome evidence is collecting; no action beyond the agreed review cadence."],
    ["Watch", "Confirm the integration owner and a dated fallback before the next checkpoint."],
    ["Blocked", "Escalate only the named decision that prevents safe progress."],
  ].forEach(([heading, body], index) => addCard(delivery, {
    name: `delivery-lane-${index + 1}`,
    left: 80 + index * 385,
    top: 186,
    width: 330,
    height: 308,
    heading,
    body,
    fill: ["#DCFCE7", "#FEF3C7", "#FEE2E2"][index],
    line: ["#86C99A", "#E8B854", "#FCA5A5"][index],
  }));
  addPresentationText(delivery, {
    name: "delivery-rule",
    text: "Rule: every watch or blocked item has one owner, one next evidence point, and one escalation decision.",
    position: { left: 82, top: 570, width: 1080, height: 70 },
    textStyle: { fontFamily: "Aptos", fontSize: 17, bold: true, color: "#1E3A8A" },
  });

  const actions = presentation.slides.add({
    name: "Decisions and owners",
    notes: "End with the decisions being asked for and the accountable next actions.",
    background: { fill: "#0F172A" },
  });
  addPresentationText(actions, {
    name: "actions-title",
    text: "Decisions and owners",
    position: { left: 80, top: 66, width: 820, height: 60 },
    textStyle: { ...titleStyle, color: "#FFFFFF" },
  });
  [
    ["Confirm", "Scope boundary", "Decision owner"],
    ["Assign", "Dependency fallback", "Delivery owner"],
    ["Review", "Outcome evidence", "Operating lead"],
  ].forEach(([verb, decision, owner], index) => actions.shapes.add({
    name: `decision-row-${index + 1}`,
    geometry: "roundRect",
    position: { left: 80, top: 184 + index * 134, width: 1110, height: 96 },
    fill: "#1E293B",
    line: { fill: "#475569", width: 1 },
    text: `${verb}: ${decision}\nOwner: ${owner}`,
    textStyle: { fontFamily: "Aptos", fontSize: 18, bold: true, color: "#F8FAFC" },
  }));
  addPresentationText(actions, {
    name: "next-review",
    text: "Next review: bring evidence, not status narration.",
    position: { left: 84, top: 612, width: 940, height: 50 },
    textStyle: { fontFamily: "Aptos", fontSize: 16, color: "#93C5FD" },
  });
  return presentation;
}

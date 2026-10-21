import type { Entity, KnowledgeGraph, Relation } from './index.js';

// The memory file historically existed in two shapes:
// 1. A single JSON document holding the whole graph: {"entities": [...], "relations": [...]}
//    (pretty-printed or compact; an empty graph could also be a bare {}).
// 2. The current canonical JSONL format: one {"type": "entity" | "relation", ...} record per line.
//
// This parser accepts both without data loss and throws on anything else so that
// corrupted files or foreign JSON shapes fail loudly instead of silently
// producing an empty graph (and then overwriting the file on the next write).

const GRAPH_KEYS = new Set(['entities', 'relations']);

const JSON_PARSE_FAILED = Symbol('jsonParseFailed');

export function parseGraphFileContent(content: string): KnowledgeGraph {
  if (content.trim() === '') {
    return { entities: [], relations: [] };
  }

  // Try to parse the whole content as a single JSON document first. This covers
  // legacy graph objects (pretty or compact) and single-line/single-record JSONL.
  const whole = tryParseJson(content);
  if (whole !== JSON_PARSE_FAILED) {
    return parseSingleDocument(whole);
  }

  // Not a single JSON document: fall back to JSONL, one record per line.
  return parseJsonlRecords(content);
}

function tryParseJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return JSON_PARSE_FAILED;
  }
}

function parseSingleDocument(value: unknown): KnowledgeGraph {
  if (!isRecord(value)) {
    throw new Error(
      `memory file must contain a JSON object graph or JSONL records, got ${describeValue(value)}`
    );
  }

  // A JSONL record (current canonical format) is tagged with a "type" field.
  if ('type' in value) {
    const graph: KnowledgeGraph = { entities: [], relations: [] };
    applyRecord(value, graph, 'JSON document');
    return graph;
  }

  // Legacy whole-graph object. Only known keys are allowed so that files with
  // foreign content (e.g. {"notes": []}) are rejected instead of silently
  // dropping data when the graph is next written back as JSONL.
  const unknownKeys = Object.keys(value).filter(key => !GRAPH_KEYS.has(key));
  if (unknownKeys.length > 0) {
    throw new Error(
      `memory file is not a valid knowledge graph: unexpected top-level key(s): ${unknownKeys.join(', ')}`
    );
  }

  return {
    entities: parseEntityArray(value.entities, 'entities'),
    relations: parseRelationArray(value.relations, 'relations'),
  };
}

function parseJsonlRecords(content: string): KnowledgeGraph {
  const graph: KnowledgeGraph = { entities: [], relations: [] };
  const lines = content.split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.trim() === '') {
      continue;
    }

    const context = `line ${i + 1}`;
    const item = tryParseJson(line);
    if (item === JSON_PARSE_FAILED) {
      throw new Error(`memory file contains invalid JSON on ${context}`);
    }
    if (!isRecord(item)) {
      throw new Error(`memory file ${context} must be a JSON object record, got ${describeValue(item)}`);
    }
    applyRecord(item, graph, context);
  }

  return graph;
}

function applyRecord(item: Record<string, unknown>, graph: KnowledgeGraph, context: string): void {
  if (item.type === 'entity') {
    graph.entities.push(parseEntity(item, context));
  } else if (item.type === 'relation') {
    graph.relations.push(parseRelation(item, context));
  } else {
    throw new Error(
      `memory file ${context} has unrecognized record type ${JSON.stringify(item.type)}` +
        ` (expected "entity" or "relation")`
    );
  }
}

function parseEntityArray(value: unknown, context: string): Entity[] {
  if (value === undefined) {
    return [];
  }
  if (!Array.isArray(value)) {
    throw new Error(`memory file "${context}" must be an array of entities`);
  }
  return value.map((item, index) => parseEntity(item, `${context}[${index}]`));
}

function parseRelationArray(value: unknown, context: string): Relation[] {
  if (value === undefined) {
    return [];
  }
  if (!Array.isArray(value)) {
    throw new Error(`memory file "${context}" must be an array of relations`);
  }
  return value.map((item, index) => parseRelation(item, `${context}[${index}]`));
}

function parseEntity(value: unknown, context: string): Entity {
  if (!isRecord(value)) {
    throw new Error(`memory file ${context} entity must be an object, got ${describeValue(value)}`);
  }
  return {
    name: requireString(value.name, 'name', context, 'entity'),
    entityType: requireString(value.entityType, 'entityType', context, 'entity'),
    observations: requireStringArray(value.observations, 'observations', context, 'entity'),
  };
}

function parseRelation(value: unknown, context: string): Relation {
  if (!isRecord(value)) {
    throw new Error(`memory file ${context} relation must be an object, got ${describeValue(value)}`);
  }
  return {
    from: requireString(value.from, 'from', context, 'relation'),
    to: requireString(value.to, 'to', context, 'relation'),
    relationType: requireString(value.relationType, 'relationType', context, 'relation'),
  };
}

function requireString(value: unknown, field: string, context: string, kind: string): string {
  if (typeof value !== 'string') {
    throw new Error(
      `memory file ${context}: ${kind} field "${field}" must be a string, got ${describeValue(value)}`
    );
  }
  return value;
}

function requireStringArray(value: unknown, field: string, context: string, kind: string): string[] {
  if (!Array.isArray(value) || value.some(item => typeof item !== 'string')) {
    throw new Error(
      `memory file ${context}: ${kind} field "${field}" must be an array of strings, got ${describeValue(value)}`
    );
  }
  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function describeValue(value: unknown): string {
  if (value === null) {
    return 'null';
  }
  if (Array.isArray(value)) {
    return 'array';
  }
  return typeof value;
}

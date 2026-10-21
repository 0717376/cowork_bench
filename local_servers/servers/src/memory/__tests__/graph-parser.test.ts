import { describe, it, expect } from 'vitest';
import { parseGraphFileContent } from '../graph-parser.js';

describe('parseGraphFileContent', () => {
  it('returns an empty graph for empty or whitespace-only content', () => {
    expect(parseGraphFileContent('')).toEqual({ entities: [], relations: [] });
    expect(parseGraphFileContent('   \n  \n')).toEqual({ entities: [], relations: [] });
  });

  it('parses a legacy graph object, pretty or compact', () => {
    const graph = {
      entities: [{ name: 'A', entityType: 'person', observations: ['x'] }],
      relations: [{ from: 'A', to: 'A', relationType: 'self' }],
    };
    expect(parseGraphFileContent(JSON.stringify(graph))).toEqual(graph);
    expect(parseGraphFileContent(JSON.stringify(graph, null, 2))).toEqual(graph);
  });

  it('parses a bare empty object and partially populated graph objects', () => {
    expect(parseGraphFileContent('{}')).toEqual({ entities: [], relations: [] });
    expect(
      parseGraphFileContent('{"entities": [{"name": "A", "entityType": "t", "observations": []}]}')
    ).toEqual({
      entities: [{ name: 'A', entityType: 't', observations: [] }],
      relations: [],
    });
  });

  it('parses single- and multi-line JSONL records', () => {
    const record = '{"type":"entity","name":"A","entityType":"t","observations":[]}';
    expect(parseGraphFileContent(record).entities).toHaveLength(1);
    expect(parseGraphFileContent(`${record}\n{"type":"relation","from":"A","to":"A","relationType":"self"}\n`).relations).toHaveLength(1);
  });

  it('throws on invalid JSON', () => {
    expect(() => parseGraphFileContent('{"entities": [')).toThrow(/invalid JSON on line 1/);
    expect(() => parseGraphFileContent('not-json')).toThrow(/invalid JSON on line 1/);
  });

  it('throws on foreign JSON shapes instead of returning an empty graph', () => {
    expect(() => parseGraphFileContent('{"notes": []}')).toThrow(/unexpected top-level key\(s\): notes/);
    expect(() => parseGraphFileContent('[]')).toThrow(/got array/);
    expect(() => parseGraphFileContent('42')).toThrow(/got number/);
    expect(() => parseGraphFileContent('null')).toThrow(/got null/);
  });

  it('throws on malformed graph members', () => {
    expect(() => parseGraphFileContent('{"entities": {}}')).toThrow(/"entities" must be an array of entities/);
    expect(() => parseGraphFileContent('{"relations": [{}]}')).toThrow(
      /relations\[0\]: relation field "from" must be a string/
    );
    expect(() =>
      parseGraphFileContent('{"entities": [{"name": "A", "entityType": "t", "observations": "no"}]}')
    ).toThrow(/entity field "observations" must be an array of strings/);
  });

  it('throws on JSONL records with unknown types', () => {
    expect(() => parseGraphFileContent('{"type": "note"}')).toThrow(/unrecognized record type "note"/);
  });
});

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { KnowledgeGraphManager } from '../index.js';

// Regression tests for loading legacy memory.json content (a single JSON
// document holding the whole graph, pretty or compact) through the real
// KnowledgeGraphManager file round-trip, alongside the canonical JSONL format.
describe('KnowledgeGraphManager memory file format compatibility', () => {
  let manager: KnowledgeGraphManager;
  let testFilePath: string;

  beforeEach(() => {
    testFilePath = path.join(
      path.dirname(fileURLToPath(import.meta.url)),
      `test-memory-format-${Date.now()}-${Math.random().toString(36).slice(2)}.json`
    );
    manager = new KnowledgeGraphManager(testFilePath);
  });

  afterEach(async () => {
    try {
      await fs.unlink(testFilePath);
    } catch {
      // Ignore errors if file doesn't exist
    }
  });

  const legacyPrettyJson = JSON.stringify(
    {
      entities: [
        {
          name: 'Recipe Selection Criteria',
          entityType: 'criteria',
          observations: ['Selected categories: Meat, Vegetable', 'Chose 10 recipes'],
        },
        {
          name: 'Category Overview',
          entityType: 'category_list',
          observations: ['Meat: Hong Shao Rou'],
        },
      ],
      relations: [
        { from: 'Recipe Selection Criteria', to: 'Category Overview', relationType: 'defines' },
      ],
    },
    null,
    2
  );

  describe('legacy JSON object graphs', () => {
    it('should read a legacy pretty-printed JSON graph without rewriting the file', async () => {
      await fs.writeFile(testFilePath, legacyPrettyJson);

      const graph = await manager.readGraph();

      expect(graph.entities).toHaveLength(2);
      expect(graph.entities[0].name).toBe('Recipe Selection Criteria');
      expect(graph.entities[0].observations).toContain('Selected categories: Meat, Vegetable');
      expect(graph.relations).toHaveLength(1);
      expect(graph.relations[0]).toEqual({
        from: 'Recipe Selection Criteria',
        to: 'Category Overview',
        relationType: 'defines',
      });

      // Reading must not rewrite the fixture
      const contentAfterRead = await fs.readFile(testFilePath, 'utf-8');
      expect(contentAfterRead).toBe(legacyPrettyJson);
    });

    it('should read a compact single-line JSON graph', async () => {
      await fs.writeFile(
        testFilePath,
        '{"entities":[{"name":"Alice","entityType":"person","observations":["likes coffee"]}],"relations":[]}'
      );

      const graph = await manager.readGraph();

      expect(graph.entities).toHaveLength(1);
      expect(graph.entities[0].name).toBe('Alice');
      expect(graph.entities[0].observations).toEqual(['likes coffee']);
      expect(graph.relations).toHaveLength(0);
    });

    it('should treat a bare empty JSON object as an empty graph', async () => {
      await fs.writeFile(testFilePath, '{}');

      const graph = await manager.readGraph();
      expect(graph).toEqual({ entities: [], relations: [] });
    });

    it('should treat an explicit empty graph object as empty', async () => {
      await fs.writeFile(testFilePath, '{"entities": [], "relations": []}');

      const graph = await manager.readGraph();
      expect(graph).toEqual({ entities: [], relations: [] });
    });

    it('should treat an empty (zero-byte) file as an empty graph', async () => {
      await fs.writeFile(testFilePath, '');

      const graph = await manager.readGraph();
      expect(graph).toEqual({ entities: [], relations: [] });
    });

    it('should preserve legacy data through create/create/delete and persist as canonical JSONL', async () => {
      await fs.writeFile(testFilePath, legacyPrettyJson);

      // Mutations must keep the legacy entities/relations and save JSONL
      const created = await manager.createEntities([
        { name: 'New Entity', entityType: 'note', observations: ['added later'] },
      ]);
      expect(created).toHaveLength(1);

      const addedRelations = await manager.createRelations([
        { from: 'New Entity', to: 'Category Overview', relationType: 'mentions' },
      ]);
      expect(addedRelations).toHaveLength(1);

      // File on disk is now canonical JSONL: one JSON record per line
      const onDisk = await fs.readFile(testFilePath, 'utf-8');
      const lines = onDisk.split('\n').filter(line => line.trim() !== '');
      expect(lines).toHaveLength(5);
      const records = lines.map(line => JSON.parse(line));
      const entityRecords = records.filter(record => record.type === 'entity');
      const relationRecords = records.filter(record => record.type === 'relation');
      expect(entityRecords).toHaveLength(3);
      expect(relationRecords).toHaveLength(2);
      expect(entityRecords.map(record => record.name)).toEqual([
        'Recipe Selection Criteria',
        'Category Overview',
        'New Entity',
      ]);

      // Reload through a fresh manager: everything persists across instances
      const manager2 = new KnowledgeGraphManager(testFilePath);
      const reloaded = await manager2.readGraph();
      expect(reloaded.entities).toHaveLength(3);
      expect(reloaded.relations).toHaveLength(2);

      // Deleting one legacy entity keeps the rest intact
      await manager2.deleteEntities(['Recipe Selection Criteria']);
      const afterDelete = await new KnowledgeGraphManager(testFilePath).readGraph();
      expect(afterDelete.entities.map(e => e.name).sort()).toEqual(['Category Overview', 'New Entity']);
      // Cascade: the relation that referenced the deleted entity is gone, the new one stays
      expect(afterDelete.relations).toEqual([
        { from: 'New Entity', to: 'Category Overview', relationType: 'mentions' },
      ]);
    });
  });

  describe('JSONL backward compatibility', () => {
    it('should read a single-line JSONL file without trailing newline', async () => {
      await fs.writeFile(
        testFilePath,
        '{"type":"entity","name":"Solo","entityType":"person","observations":[]}'
      );

      const graph = await manager.readGraph();
      expect(graph.entities).toEqual([{ name: 'Solo', entityType: 'person', observations: [] }]);
      expect(graph.relations).toHaveLength(0);
    });

    it('should read a multi-line JSONL file with entities and relations', async () => {
      await fs.writeFile(
        testFilePath,
        [
          '{"type":"entity","name":"Alice","entityType":"person","observations":["a"]}',
          '{"type":"entity","name":"Bob","entityType":"person","observations":[]}',
          '{"type":"relation","from":"Alice","to":"Bob","relationType":"knows"}',
          '',
        ].join('\n')
      );

      const graph = await manager.readGraph();
      expect(graph.entities).toHaveLength(2);
      expect(graph.relations).toEqual([{ from: 'Alice', to: 'Bob', relationType: 'knows' }]);

      // Appending via the manager preserves the existing JSONL records
      await manager.createEntities([
        { name: 'Carol', entityType: 'person', observations: [] },
      ]);
      const reloaded = await new KnowledgeGraphManager(testFilePath).readGraph();
      expect(reloaded.entities.map(e => e.name)).toEqual(['Alice', 'Bob', 'Carol']);
      expect(reloaded.relations).toHaveLength(1);
    });
  });

  describe('invalid memory files fail loudly', () => {
    const invalidCases: Array<{ name: string; content: string; message: RegExp }> = [
      {
        name: 'malformed pretty JSON',
        content: '{\n  "entities": [\n    {"name": ',
        message: /Failed to parse memory file .*: memory file contains invalid JSON/,
      },
      {
        name: 'foreign JSON shape with unexpected top-level key',
        content: '{"notes": []}',
        message: /unexpected top-level key\(s\): notes/,
      },
      {
        name: 'graph with non-array entities',
        content: '{"entities": "not-an-array"}',
        message: /"entities" must be an array of entities/,
      },
      {
        name: 'graph with malformed entity member',
        content: '{"entities": [{"name": 42, "entityType": "person", "observations": []}]}',
        message: /entities\[0\]: entity field "name" must be a string, got number/,
      },
      {
        name: 'JSON array root',
        content: '[{"type":"entity"}]',
        message: /must contain a JSON object graph or JSONL records, got array/,
      },
      {
        name: 'JSONL with a garbage line',
        content:
          '{"type":"entity","name":"A","entityType":"person","observations":[]}\nthis-is-not-json',
        message: /invalid JSON on line 2/,
      },
      {
        name: 'JSONL line that is not an entity/relation record',
        content:
          '{"type":"entity","name":"A","entityType":"person","observations":[]}\n{"type":"note","payload":"hi"}',
        message: /line 2 has unrecognized record type "note"/,
      },
    ];

    for (const { name, content, message } of invalidCases) {
      it(`should reject ${name} instead of returning an empty graph`, async () => {
        await fs.writeFile(testFilePath, content);

        await expect(manager.readGraph()).rejects.toThrow(message);

        // A failed read must not modify or remove the file
        const contentAfter = await fs.readFile(testFilePath, 'utf-8');
        expect(contentAfter).toBe(content);
      });
    }
  });
});

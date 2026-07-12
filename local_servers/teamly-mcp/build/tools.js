import { z } from 'zod';
import { query, getSpaceByKey, findPage, getLabels } from './db.js';
function asText(payload) {
    return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}
function asError(msg, extra) {
    return asText({ error: msg, ...(extra ?? {}) });
}
export function registerTools(server) {
    // ---------------------------------------------------------------- spaces
    server.tool('list_spaces', 'Возвращает список всех пространств Teamly (ключ, название, описание).', {}, async () => {
        const r = await query('SELECT key, name, description FROM teamly.spaces ORDER BY key');
        return asText({ spaces: r.rows });
    });
    server.tool('create_space', 'Создаёт новое пространство Teamly. Если ключ уже существует — возвращает ошибку.', {
        key: z.string().min(1).describe('Уникальный короткий ключ пространства, например "TRIPS".'),
        name: z.string().min(1).describe('Человекочитаемое название пространства.'),
        description: z.string().optional().describe('Описание (необязательно).'),
    }, async ({ key, name, description }) => {
        const exists = await getSpaceByKey(key);
        if (exists)
            return asError('Пространство с таким ключом уже существует', { key });
        const r = await query('INSERT INTO teamly.spaces (key, name, description) VALUES ($1, $2, $3) RETURNING id, key, name, description', [key, name, description ?? '']);
        return asText({ space: r.rows[0] });
    });
    // ---------------------------------------------------------------- pages
    server.tool('list_pages', 'Список страниц. Можно отфильтровать по space_key (рекомендуется). Возвращает краткое представление без тела.', {
        space_key: z.string().optional().describe('Ключ пространства для фильтра.'),
        limit: z.number().int().min(1).max(500).optional().describe('Макс. число страниц (по умолчанию 100).'),
    }, async ({ space_key, limit }) => {
        const lim = limit ?? 100;
        const sql = `
        SELECT p.id, s.key AS space_key, p.title, p.author, p.parent_id,
               p.updated_at::text AS updated_at
          FROM teamly.pages p
          JOIN teamly.spaces s ON s.id = p.space_id
         ${space_key ? 'WHERE s.key = $1' : ''}
         ORDER BY p.updated_at DESC
         LIMIT ${lim}`;
        const params = space_key ? [space_key] : [];
        const r = await query(sql, params);
        return asText({ pages: r.rows, count: r.rowCount });
    });
    server.tool('get_page', 'Получить страницу по id или по комбинации (space_key + title). Возвращает полный текст и метки.', {
        page_id: z.number().int().optional().describe('Числовой id страницы.'),
        space_key: z.string().optional().describe('Ключ пространства (если ищем по заголовку).'),
        title: z.string().optional().describe('Заголовок страницы (если ищем без id).'),
    }, async ({ page_id, space_key, title }) => {
        let page = null;
        if (page_id !== undefined) {
            const r = await query(`SELECT p.id, p.space_id, s.key AS space_key, p.title, p.body, p.author,
                  p.parent_id, p.created_at::text, p.updated_at::text
             FROM teamly.pages p JOIN teamly.spaces s ON s.id = p.space_id
            WHERE p.id = $1`, [page_id]);
            page = r.rows[0] ?? null;
        }
        else if (title) {
            page = await findPage(space_key, title);
        }
        else {
            return asError('Нужно передать либо page_id, либо title (и опционально space_key).');
        }
        if (!page)
            return asError('Страница не найдена', { page_id, space_key, title });
        const labels = await getLabels(page.id);
        return asText({ page: { ...page, labels } });
    });
    server.tool('create_page', 'Создаёт страницу в указанном пространстве. Возвращает созданную страницу с id.', {
        space_key: z.string().describe('Ключ пространства (например "TRIPS").'),
        title: z.string().min(1).describe('Заголовок страницы.'),
        body: z.string().describe('Содержимое страницы (markdown).'),
        author: z.string().optional().describe('Имя автора (необязательно).'),
        parent_title: z.string().optional().describe('Заголовок родительской страницы в том же пространстве (необязательно).'),
        labels: z.array(z.string()).optional().describe('Метки для страницы (создаются автоматически).'),
    }, async ({ space_key, title, body, author, parent_title, labels }) => {
        const space = await getSpaceByKey(space_key);
        if (!space)
            return asError('Пространство не найдено', { space_key });
        let parent_id = null;
        if (parent_title) {
            const parent = await findPage(space_key, parent_title);
            if (!parent)
                return asError('Родительская страница не найдена', { parent_title, space_key });
            parent_id = parent.id;
        }
        const r = await query(`INSERT INTO teamly.pages (space_id, title, body, author, parent_id)
         VALUES ($1, $2, $3, $4, $5) RETURNING id`, [space.id, title, body, author ?? '', parent_id]);
        const pageId = r.rows[0].id;
        if (labels && labels.length) {
            for (const name of labels) {
                await query('INSERT INTO teamly.labels (name) VALUES ($1) ON CONFLICT (name) DO NOTHING', [name]);
                await query(`INSERT INTO teamly.page_labels (page_id, label_id)
             SELECT $1, id FROM teamly.labels WHERE name = $2
             ON CONFLICT DO NOTHING`, [pageId, name]);
            }
        }
        const full = await query(`SELECT p.id, p.space_id, s.key AS space_key, p.title, p.body, p.author,
                p.parent_id, p.created_at::text, p.updated_at::text
           FROM teamly.pages p JOIN teamly.spaces s ON s.id = p.space_id
          WHERE p.id = $1`, [pageId]);
        const lbls = await getLabels(pageId);
        return asText({ page: { ...full.rows[0], labels: lbls } });
    });
    server.tool('update_page', 'Обновляет тело и/или заголовок страницы. Можно искать страницу по id или (space_key+title).', {
        page_id: z.number().int().optional(),
        space_key: z.string().optional(),
        title: z.string().optional(),
        new_title: z.string().optional().describe('Новый заголовок (если меняем).'),
        body: z.string().optional().describe('Новое содержимое (если меняем).'),
    }, async ({ page_id, space_key, title, new_title, body }) => {
        let target = null;
        if (page_id !== undefined) {
            const r = await query('SELECT * FROM teamly.pages WHERE id = $1', [page_id]);
            target = r.rows[0] ?? null;
        }
        else if (title) {
            target = await findPage(space_key, title);
        }
        else {
            return asError('Нужно передать page_id или title.');
        }
        if (!target)
            return asError('Страница не найдена');
        const sets = [];
        const params = [];
        let i = 1;
        if (new_title !== undefined) {
            sets.push(`title = $${i++}`);
            params.push(new_title);
        }
        if (body !== undefined) {
            sets.push(`body = $${i++}`);
            params.push(body);
        }
        if (!sets.length)
            return asError('Нечего обновлять — передайте new_title и/или body.');
        sets.push(`updated_at = now()`);
        params.push(target.id);
        await query(`UPDATE teamly.pages SET ${sets.join(', ')} WHERE id = $${i}`, params);
        const r = await query(`SELECT p.id, p.space_id, s.key AS space_key, p.title, p.body, p.author,
                p.parent_id, p.created_at::text, p.updated_at::text
           FROM teamly.pages p JOIN teamly.spaces s ON s.id = p.space_id
          WHERE p.id = $1`, [target.id]);
        const labels = await getLabels(target.id);
        return asText({ page: { ...r.rows[0], labels } });
    });
    server.tool('search_pages', 'Полнотекстовый поиск по заголовку и телу страниц (ILIKE). Опц. фильтр по space_key.', {
        query: z.string().min(1).describe('Поисковая строка.'),
        space_key: z.string().optional(),
        limit: z.number().int().min(1).max(100).optional(),
    }, async ({ query: q, space_key, limit }) => {
        const lim = limit ?? 25;
        const sql = `
        SELECT p.id, s.key AS space_key, p.title, p.author, p.updated_at::text AS updated_at,
               LEFT(p.body, 200) AS preview
          FROM teamly.pages p
          JOIN teamly.spaces s ON s.id = p.space_id
         WHERE (p.title ILIKE $1 OR p.body ILIKE $1)
           ${space_key ? 'AND s.key = $2' : ''}
         ORDER BY p.updated_at DESC
         LIMIT ${lim}`;
        const params = space_key ? [`%${q}%`, space_key] : [`%${q}%`];
        const r = await query(sql, params);
        return asText({ results: r.rows, count: r.rowCount });
    });
    server.tool('add_label', 'Добавить метку к странице (метка создаётся, если её нет).', {
        page_id: z.number().int(),
        label: z.string().min(1),
    }, async ({ page_id, label }) => {
        const exists = await query('SELECT 1 FROM teamly.pages WHERE id = $1', [page_id]);
        if (!exists.rowCount)
            return asError('Страница не найдена', { page_id });
        await query('INSERT INTO teamly.labels (name) VALUES ($1) ON CONFLICT (name) DO NOTHING', [label]);
        await query(`INSERT INTO teamly.page_labels (page_id, label_id)
         SELECT $1, id FROM teamly.labels WHERE name = $2
         ON CONFLICT DO NOTHING`, [page_id, label]);
        const labels = await getLabels(page_id);
        return asText({ page_id, labels });
    });
    server.tool('delete_page', 'Удаляет страницу по id. Дочерние страницы (если есть) теряют ссылку на родителя.', {
        page_id: z.number().int(),
    }, async ({ page_id }) => {
        const r = await query('DELETE FROM teamly.pages WHERE id = $1 RETURNING id', [page_id]);
        if (!r.rowCount)
            return asError('Страница не найдена', { page_id });
        return asText({ deleted: page_id });
    });
}

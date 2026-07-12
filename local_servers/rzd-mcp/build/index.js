#!/usr/bin/env node
// PostgreSQL-backed RZD (РЖД) MCP server. Mirrors the shape of 12306-mcp but
// queries the rzd.* schema and exposes Russian-language tool descriptions.
import { program } from 'commander';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { format } from 'date-fns';
import { toZonedTime } from 'date-fns-tz';
import pg from 'pg';
const { Pool } = pg;
const VERSION = '0.1.0';
const pool = new Pool({
    host: process.env.PG_HOST || 'localhost',
    port: parseInt(process.env.PG_PORT || '5432'),
    database: process.env.PG_DATABASE || 'cowork_gym',
    user: process.env.PG_USER || 'postgres',
    password: process.env.PG_PASSWORD || 'postgres',
    idleTimeoutMillis: 10000,
});
pool.on('error', () => { });
// Station dictionaries (loaded once from rzd.stations at startup)
let STATIONS = {};
let CITY_STATIONS = {};
let CITY_CODES = {};
let NAME_STATIONS = {};
async function loadStations() {
    const result = await pool.query(`SELECT station_code, station_name, station_short, city FROM rzd.stations`);
    STATIONS = {};
    CITY_STATIONS = {};
    CITY_CODES = {};
    NAME_STATIONS = {};
    for (const row of result.rows) {
        const s = row;
        STATIONS[s.station_code] = s;
        if (!CITY_STATIONS[s.city])
            CITY_STATIONS[s.city] = [];
        CITY_STATIONS[s.city].push({ station_code: s.station_code, station_name: s.station_name });
        if (s.station_name === s.city || !CITY_CODES[s.city]) {
            CITY_CODES[s.city] = { station_code: s.station_code, station_name: s.station_name };
        }
        NAME_STATIONS[s.station_name] = { station_code: s.station_code, station_name: s.station_name };
    }
}
function formatAvailability(num) {
    if (num.match(/^\d+$/)) {
        const n = parseInt(num);
        return n === 0 ? 'Мест нет' : `Свободно ${n} мест`;
    }
    switch (num) {
        case 'много':
        case 'есть':
            return 'Места есть';
        case 'нет':
        case '--':
        case '':
            return 'Мест нет';
        default:
            return num;
    }
}
function formatTicketsInfo(tickets) {
    if (tickets.length === 0)
        return 'По заданному маршруту поездов не найдено.';
    let out = 'Поезд | Откуда → Куда | Отправление → Прибытие | В пути\n';
    for (const t of tickets) {
        let s = `${t.start_train_code} (train_no: ${t.train_no}) `
            + `${t.from_station} (telecode: ${t.from_station_telecode}) → `
            + `${t.to_station} (telecode: ${t.to_station_telecode}) `
            + `${t.start_time} → ${t.arrive_time}, в пути: ${t.lishi}`;
        for (const p of t.prices) {
            s += `\n- ${p.seat_name}: ${formatAvailability(p.num)}, цена ${p.price}₽`;
        }
        out += `${s}\n`;
    }
    return out;
}
const TIME_COMPARATORS = {
    startTime: (a, b) => toMin(a.start_time) - toMin(b.start_time),
    arriveTime: (a, b) => toMin(a.arrive_time) - toMin(b.arrive_time),
    duration: (a, b) => toMin(a.lishi) - toMin(b.lishi),
};
function toMin(hhmm) {
    const [h, m] = hhmm.split(':').map(Number);
    return h * 60 + (m || 0);
}
function filterTickets(tickets, earliestStartTime = 0, latestStartTime = 24, sortFlag = '', sortReverse = false, limitedNum = 0) {
    let result = tickets.filter(t => {
        const h = parseInt(t.start_time.split(':')[0], 10);
        return h >= earliestStartTime && h < latestStartTime;
    });
    if (sortFlag && TIME_COMPARATORS[sortFlag]) {
        result.sort(TIME_COMPARATORS[sortFlag]);
        if (sortReverse)
            result.reverse();
    }
    return limitedNum > 0 ? result.slice(0, limitedNum) : result;
}
async function queryTickets(date, fromStation, toStation) {
    const result = await pool.query(`SELECT t.id, t.train_no, t.station_train_code, t.from_station_telecode, t.to_station_telecode,
                t.start_time, t.arrive_time, t.lishi,
                sf.station_name AS from_name, st.station_name AS to_name
         FROM rzd.trains t
         JOIN rzd.stations sf ON sf.station_code = t.from_station_telecode
         JOIN rzd.stations st ON st.station_code = t.to_station_telecode
         WHERE t.from_station_telecode = $1
           AND t.to_station_telecode = $2
           AND t.depart_date = $3`, [fromStation, toStation, date]);
    const tickets = [];
    for (const row of result.rows) {
        const seats = await pool.query(`SELECT seat_type_code, seat_name, seat_short, num, price
             FROM rzd.train_seats WHERE train_id = $1`, [row.id]);
        const prices = seats.rows.map(s => ({
            seat_name: s.seat_name,
            short: s.seat_short,
            seat_type_code: s.seat_type_code,
            num: s.num,
            price: parseFloat(s.price),
        }));
        tickets.push({
            train_no: row.train_no,
            start_train_code: row.station_train_code,
            start_date: date,
            arrive_date: date,
            start_time: row.start_time,
            arrive_time: row.arrive_time,
            lishi: row.lishi,
            from_station: row.from_name,
            to_station: row.to_name,
            from_station_telecode: row.from_station_telecode,
            to_station_telecode: row.to_station_telecode,
            prices,
        });
    }
    return tickets;
}
export const server = new McpServer({
    name: 'rzd-mcp',
    version: VERSION,
});
server.resource('stations', 'data://all-stations', async (uri) => ({
    contents: [{ uri: uri.href, text: JSON.stringify(STATIONS) }],
}));
server.tool('get-current-date', 'Возвращает текущую дату в часовом поясе Москвы (Europe/Moscow, UTC+3) в формате "yyyy-MM-dd". Используй для разрешения относительных дат ("завтра", "в следующий понедельник") перед вызовом get-trains.', {}, async () => {
    const moscow = toZonedTime(new Date(), 'Europe/Moscow');
    return { content: [{ type: 'text', text: format(moscow, 'yyyy-MM-dd') }] };
});
server.tool('get-stations-in-city', 'По названию города возвращает список всех вокзалов с их кодами (station_code). Пример: для "Москва" вернёт Москву-Ленинградскую, Москву-Курскую и т.д.', { city: z.string().describe('Название города на русском, например "Москва" или "Санкт-Петербург".') }, async ({ city }) => {
    if (!(city in CITY_STATIONS)) {
        return { content: [{ type: 'text', text: 'Ошибка: город не найден.' }] };
    }
    return { content: [{ type: 'text', text: JSON.stringify(CITY_STATIONS[city]) }] };
});
server.tool('get-station-code-by-city', 'По названию города возвращает основной (по умолчанию) station_code этого города. Для нескольких городов разделяй их символом "|" (например "Москва|Санкт-Петербург"). Используй когда пользователь называет только город, без конкретного вокзала.', { cities: z.string().describe('Город(а) на русском. Несколько — через "|".') }, async ({ cities }) => {
    const result = {};
    for (const city of cities.split('|')) {
        result[city] = city in CITY_CODES ? CITY_CODES[city] : { error: 'Город не найден.' };
    }
    return { content: [{ type: 'text', text: JSON.stringify(result) }] };
});
server.tool('get-station-code-by-name', 'По точному названию конкретного вокзала возвращает его station_code. Несколько названий — через "|" (например "Москва-Ленинградская|Санкт-Петербург-Главный"). Используй когда пользователь называет конкретный вокзал.', { stationNames: z.string().describe('Названия вокзалов на русском, через "|".') }, async ({ stationNames }) => {
    const result = {};
    for (const name of stationNames.split('|')) {
        result[name] = name in NAME_STATIONS ? NAME_STATIONS[name] : { error: 'Вокзал не найден.' };
    }
    return { content: [{ type: 'text', text: JSON.stringify(result) }] };
});
server.tool('get-station-by-code', 'По station_code (3-буквенный код, например "MOW") возвращает полную информацию о вокзале: имя, город. Используется для обратного поиска по коду.', { stationCode: z.string().describe('3-буквенный station_code, например "MOW" или "SPB".') }, async ({ stationCode }) => {
    if (!STATIONS[stationCode])
        return { content: [{ type: 'text', text: 'Ошибка: вокзал не найден.' }] };
    return { content: [{ type: 'text', text: JSON.stringify(STATIONS[stationCode]) }] };
});
server.tool('get-trains', 'Поиск поездов РЖД на заданную дату между двумя станциями. Возвращает список поездов с временем отправления/прибытия, длительностью в пути и наличием/ценой по классам (Эконом, Эконом+, Бизнес).', {
    date: z.string().length(10).describe('Дата в формате yyyy-MM-dd. Для относительных дат сперва вызови get-current-date.'),
    fromStation: z.string().describe('station_code станции отправления (например "MOW"). НЕ передавай сюда название кириллицей — сперва получи код через get-station-code-by-name или get-station-code-by-city.'),
    toStation: z.string().describe('station_code станции прибытия. Те же правила, что для fromStation.'),
    earliestStartTime: z.number().min(0).max(24).optional().default(0).describe('Самое раннее время отправления (час, 0-24). По умолчанию 0.'),
    latestStartTime: z.number().min(0).max(24).optional().default(24).describe('Самое позднее время отправления (час, 0-24). По умолчанию 24.'),
    sortFlag: z.string().optional().default('').describe('Сортировка: startTime (по отправлению), arriveTime (по прибытию), duration (по длительности). Пусто — без сортировки.'),
    sortReverse: z.boolean().optional().default(false).describe('Инвертировать порядок сортировки. Работает только вместе с sortFlag.'),
    limitedNum: z.number().min(0).optional().default(0).describe('Лимит количества поездов в ответе. 0 — без лимита.'),
}, async ({ date, fromStation, toStation, earliestStartTime, latestStartTime, sortFlag, sortReverse, limitedNum }) => {
    if (!STATIONS[fromStation] || !STATIONS[toStation]) {
        return { content: [{ type: 'text', text: 'Ошибка: одна из станций не найдена. Проверь station_code.' }] };
    }
    const tickets = await queryTickets(date, fromStation, toStation);
    const filtered = filterTickets(tickets, earliestStartTime, latestStartTime, sortFlag || '', sortReverse, limitedNum);
    return { content: [{ type: 'text', text: formatTicketsInfo(filtered) }] };
});
server.tool('get-train-route', 'Возвращает полный маршрут поезда: список промежуточных станций с временем прибытия, отправления и длительностью стоянки. Используй когда нужно знать, через какие станции идёт конкретный поезд.', {
    trainNo: z.string().describe('Поле train_no конкретного поезда (берётся из ответа get-trains, например "752A_260310_1").'),
}, async ({ trainNo }) => {
    const result = await pool.query(`SELECT station_no, station_telecode, station_name, arrive_time, depart_time, stopover_time
             FROM rzd.train_routes WHERE train_no = $1 ORDER BY station_no`, [trainNo]);
    if (result.rows.length === 0)
        return { content: [{ type: 'text', text: 'Маршрут поезда не найден.' }] };
    const route = result.rows.map(r => ({
        station_no: r.station_no,
        station_name: r.station_name,
        arrive_time: r.arrive_time,
        depart_time: r.depart_time,
        stopover_time: r.stopover_time,
    }));
    return { content: [{ type: 'text', text: JSON.stringify(route) }] };
});
async function startServer() {
    await loadStations();
    console.error(`RZD MCP Server (pg-backed) запущен. Загружено ${Object.keys(STATIONS).length} вокзалов.`);
    program
        .name('mcp-server-rzd')
        .version(VERSION)
        .option('--stdio', 'use stdio transport (default)', true)
        .parse(process.argv);
    const transport = new StdioServerTransport();
    await server.connect(transport);
}
startServer().catch(err => {
    console.error('Не удалось запустить RZD MCP:', err);
    process.exit(1);
});

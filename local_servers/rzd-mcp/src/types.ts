export type TicketInfo = {
    train_no: string;
    start_train_code: string;
    start_date: string;
    start_time: string;
    arrive_date: string;
    arrive_time: string;
    lishi: string;            // travel duration, "HH:MM"
    from_station: string;
    to_station: string;
    from_station_telecode: string;
    to_station_telecode: string;
    prices: Price[];
};

export type StationData = {
    station_code: string;
    station_name: string;
    station_short: string;
    city: string;
};

export interface Price {
    seat_name: string;        // e.g. "Эконом", "Эконом+", "Бизнес"
    short: string;            // e.g. "ec", "ec+", "biz"
    seat_type_code: string;
    num: string;              // availability marker, e.g. "12", "0", "много"
    price: number;            // RUB
}

export type RouteStationInfo = {
    arrive_time: string;
    depart_time: string;
    station_name: string;
    stopover_time: string;    // minutes string, e.g. "2"
    station_no: number;
};

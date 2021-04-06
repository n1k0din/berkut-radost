import argparse
from datetime import datetime, timedelta
from random import choice

from docxtpl import DocxTemplate


INTERVAL_MIN = timedelta(minutes=90)
INTERVAL_MAX = timedelta(minutes=120)
MINUTES_STEP = timedelta(minutes=5)
MINUTES_TO_FIRST_CRAWL = timedelta(minutes=10)
ONE_DAY = timedelta(days=1)


def create_arguments_parser():
    parser = argparse.ArgumentParser(description='Создает doc-файлы с расписанием обходов')
    parser.add_argument(
        '-d',
        '--date',
        help='На какой день создать расписание в виде гггг-мм-дд, по-умолчанию на завтра',
        type=datetime.fromisoformat,
        default=tommorow(),
    )

    parser.add_argument(
        '-n',
        '--num',
        help='Количество смен, по-умолчанию 1',
        type=int,
        default=1,
    )

    parser.add_argument(
        '-s',
        '--start',
        help='Час начала смены, по-умолчанию 8',
        type=int,
        default=8,
    )

    return parser


def get_args_from_parser(parser):
    args = parser.parse_args()

    return args.date, args.num, args.start


def tommorow():
    tommorow = datetime.today() + ONE_DAY
    return tommorow.replace(hour=0, minute=0, second=0, microsecond=0)


def generate_intervals():
    intervals = []
    interval = INTERVAL_MIN
    while interval <= INTERVAL_MAX:
        intervals.append(interval)
        interval = interval + MINUTES_STEP

    return intervals


def generate_crawl_schedule(shift_start_time, shift_stop_time, intervals):
    crawls = []
    current_crawl = shift_start_time + MINUTES_TO_FIRST_CRAWL
    while current_crawl < shift_stop_time:
        crawls.append(current_crawl)
        current_crawl += choice(intervals)

    return crawls


def get_formatted_crawls(crawls):
    formatted_crawls = []
    for crawl in crawls:
        formatted_crawls.append(crawl.strftime("%H:%M %d.%m.%Y"))

    return formatted_crawls


def get_objects_from_file(filename='objects.txt'):
    with open(filename, 'r') as f:
        return [object_name.strip() for object_name in f]


def build_objects(objects_names, shift_start, intervals):
    objects = []
    for num, object_name in enumerate(objects_names, start=1):
        object = {'num': num, 'name': object_name}
        crawls = generate_crawl_schedule(
            shift_start,
            shift_start + ONE_DAY,
            intervals
        )
        formatted_crawls = get_formatted_crawls(crawls)
        object['crawls'] = formatted_crawls
        objects.append(object)

    return objects


def build_context(objects, shift_start):
    return {
        'shift_start_day': shift_start.strftime('%d.%m.%Y'),
        'shift_end_day': (shift_start + ONE_DAY).strftime('%d.%m.%Y'),
        'objects': objects,
    }


def main():
    arg_parser = create_arguments_parser()
    shift_date, shifts_amount, start_hour = get_args_from_parser(arg_parser)

    shift_start_daytime = shift_date.replace(hour=start_hour)

    objects_names = get_objects_from_file()

    intervals = generate_intervals()

    shift_start = shift_start_daytime

    for _ in range(shifts_amount):
        doc = DocxTemplate('template.docx')

        objects = build_objects(objects_names, shift_start, intervals)

        context = build_context(objects, shift_start)

        doc.render(context)
        filename = f'output/{shift_start.strftime("%d%m%Y")}-'\
            f'{(shift_start + ONE_DAY).strftime("%d%m%Y")}.docx'
        doc.save(filename)

        shift_start += ONE_DAY


if __name__ == '__main__':
    main()

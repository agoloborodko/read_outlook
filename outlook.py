from inspect import getmembers
import pandas as pd
import dateutil.parser


def print_members(obj, obj_name="placeholder_name"):
    """Print members of given COM object"""
    try:
        fields = list(obj._prop_map_get_.keys())
    except AttributeError:
        print("Object has no attribute '_prop_map_get_'")
        print("Check if the initial COM object was created with"
              "'win32com.client.gencache.EnsureDispatch()'")
        raise
    methods = [m[0] for m in getmembers(obj) if (not m[0].startswith("_")
                                                 and "clsid" not in m[0].lower())]

    if len(fields) + len(methods) > 0:
        print("Members of '{}' ({}):".format(obj_name, obj))
    else:
        raise ValueError("Object has no members to print")

    print("\tFields:")
    if fields:
        for field in fields:
            print(f"\t\t{field}")
    else:
        print("\t\tObject has no fields to print")

    print("\tMethods:")
    if methods:
        for method in methods:
            print(f"\t\t{method}")
    else:
        print("\t\tObject has no methods to print")


def outlook_get_folders_list(outlook_obj):
    inbox = outlook_obj.GetDefaultFolder(6)
    for folder in inbox.Folders:
        print(folder.Name)


def outlook_get_folder_from_name(outlook_obj, folder_name):
    inbox = outlook_obj.GetDefaultFolder(6)
    for folder in inbox.Folders:
        if folder.Name == folder_name:
            return folder
    return None


def get_recieved_time(item):
    return pd.to_datetime(dateutil.parser.parse(str(item.ReceivedTime))).tz_convert(tz=None)


def get_sender_email(item):
    return str(item.SenderEmailAddress)


def date_correction(index, target_date, items, border='upper'):
    interval = get_recieved_time(items[items.Count-1]) - get_recieved_time(items[0])
    items_count = items.Count

    while True:
        current_date = get_recieved_time(items[index])
        correction = int(((current_date - target_date) / interval) * items_count)
        corr_date = get_recieved_time(items[index - correction])

        print('trying index {}, correction {:.1%}, got {}'.format(
            index - correction, -correction / items.Count, corr_date)
        )

        if abs(correction) < 1:
            if correction < 0:
                correction = -1
            else:
                correction = 1

        if corr_date < current_date < target_date or corr_date > current_date > target_date:  # удаляемся от таргета
            # уменьшаем абсолютное значение correction
            return date_correction(index + int(correction / 2) + 1, target_date, items, border)
        if current_date < target_date < corr_date or current_date > target_date > corr_date:  # перескочили таргет
            # перезадаем границы
            interval = abs(corr_date - current_date)
            items_count = abs(correction)
        else:  # приближаемся к таргету
            index -= correction
            if index < 1:
                index = 1

        #  Проверяем, не достигли ли мы таргета
        if current_date == target_date:
            return index
        if current_date > target_date:
            if index == 1:
                return index
            if get_recieved_time(items[index - 1]) < target_date:
                if border == 'upper':
                    return index - 1
                elif border == 'lower':
                    return index
        elif current_date < target_date:
            if index == items.Count:
                return index
            if get_recieved_time(items[index + 1]) > target_date:
                if border == 'upper':
                    return index
                elif border == 'lower':
                    return index + 1


def find_msg_by_date(items, date_start, date_end):
    msgs_count = items.Count - 1
    date_first_msg = get_recieved_time(items[0])
    date_last_msg = get_recieved_time(items[msgs_count])

    if date_last_msg < date_end:
        date_end = date_last_msg

    whole_interval = date_last_msg - date_first_msg
    target_interval = date_end - date_start
    target_interval_border = date_last_msg - date_end

    date_end_location = int(msgs_count - ((target_interval_border / whole_interval) * msgs_count))
    date_start_location = int(date_end_location - ((target_interval / whole_interval) * msgs_count))
    print('whole interval')
    print(date_first_msg, date_last_msg, '\n')
    print('desired interval')
    print(date_start, date_end)
    print(date_start_location, date_end_location, '\n')
    print('search start interval')
    print(get_recieved_time(items[date_start_location]), get_recieved_time(items[date_end_location]))
    print(date_start_location, date_end_location, '\n')

    print('correcting date_end')

    date_end_location = date_correction(
        index=date_end_location,
        target_date=date_end,
        items=items
    )

    print('correcting date_start')

    date_start_location = date_correction(
        index=date_start_location,
        target_date=date_start,
        items=items
    )

    print('search result interval')
    print(get_recieved_time(items[date_start_location]), get_recieved_time(items[date_end_location]), '\n')

    return date_start_location, date_end_location


def extract_msg_by_dates(outlook_obj, folder_name, date_start, date_end):
    f = outlook_get_folder_from_name(outlook_obj, folder_name)
    f_items = f.Items
    f_items.Sort("[ReceivedTime]")

    start, end = find_msg_by_date(
        f_items, date_start, date_end
    )

    result = []

    for i in range(start, end):
        if (i - start) % 50 == 0:
            print('Загружаю {} из {}'.format(i - start, end - start))
        result.append(
            [
                i,
                get_recieved_time(f_items[i]),
                f_items[i].HTMLBody,
                get_sender_email(f_items[i])
            ]
        )

    return result

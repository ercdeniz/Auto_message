import os

import utils


def main():
    logs_directory = "../logs"
    if not os.path.exists(logs_directory):
        os.mkdir(logs_directory)

    people = utils.get_data("../data/people_info.xlsx", "Sayfa1")
    day_left = utils.day_calculator(2024, 7, 7)

    for index in range(len(people)):
        try:
            # reminder, control, request
            message_type = "control"
            utils.send_message(people, index, day_left, message_type)
        except Exception as exception:
            utils.logger(people, exception, os.path.join(logs_directory, "error_log.txt"))

    utils.move_file_with_timestamp(logs_directory)


if __name__ == "__main__":
    main()

import functions as f
import logging
import inspect
import datetime


def get_initials(geologist_name): # helper function to add_initials()
    for i, letter in enumerate(geologist_name):
        if letter == ' ':
            break
    return geologist_name[0] + geologist_name[i + 1]


def add_initials(orig, s, geologist_name, FIELD_IDs):
    try:
        if orig[-3].isalpha:
            return False, s
    except Exception as e:
        return False, s

    count = 0
    for letter in s:
        if letter.isalpha():
            count += 1
        else:
            break

    try:
        initials = get_initials(geologist_name)
    except:
        return False, s

    new_code = s[:count] + initials + s[count:]  # try appending after the first set of initials
    if find_location(orig, new_code, FIELD_IDs)[0]:
        return (True, new_code)

    new_code = initials + s[:count] + s[count:] # try appending initials to the front
    if find_location(orig, new_code, FIELD_IDs)[0]:
        return (True, new_code)

    return (False, new_code)


def remove_trailing_letters(orig, s, FIELD_IDs):
    transitioned = False
    for i, letter in enumerate(s):
        if letter.isalpha():
            if transitioned:
                if find_location(orig, s[:i], FIELD_IDs)[0]:
                    return (True, s[:i])
                return (False, s[:i])
        else:
            transitioned = True # it is a number now
    return False, s


def add_prefix_number(orig, s, FIELD_IDs):
    try:
        if not orig[2].isalpha():
            return (False, s)
    except:
        return False, s

    for i, letter in enumerate(s):
        if not letter.isalpha():
            break

    new_code = s[:i] + '0' + s[i:]
    if find_location(orig, new_code, FIELD_IDs)[0]:
        return (True, new_code)
    return (False, new_code)


def find_location(orig, s, FIELD_IDs):
    if s in FIELD_IDs:
        print(s, 'found!')
        logging.info(orig + '\t' + s + '\t\tfound! \tmethod: {}'.format(inspect.stack()[1].function if inspect.stack()[1].function != '<module>' else 'exact'))
        # input('Press enter to continue ...')
        return (True, s)
    return False, s


def not_found(orig, BLOCK_NUMBER):
    print('Out of options')
    print('did not find for {}'.format(orig))
    logging.info(orig + '\t' + BLOCK_NUMBER + '\t\tno match!')


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG,
                        filename='app.log',
                        filemode='w',
                        format='%(name)s - %(levelname)s - %(message)s'
                        )

    # load the excel docs
    gsq = f.load_workbook('gsq.xlsx')
    merlin = f.load_workbook('merlin.xlsx')

    # grab the current worksheets
    ws = gsq.active
    merlin_ws = merlin.active

    # manually create Merlin data's dimensions
    merlin_dimensions = f.Dimension('A', '1', 'G', '184899')

    dimensions = f.get_dimensions(ws)  # dimensions of the gsq excel doc

    # get the FIELD_ID cells from the merlin excel
    # FIELD_IDs = [f for f in merlin_ws['G'] if f.coordinate != 'G1']
    FIELD_IDs = dict([(f.value, f) for f in merlin_ws['G'] if f.coordinate != 'G1'])

    row_count = 0
    for row in ws.iter_rows(min_row=dimensions.min_row + 1, max_row=dimensions.max_row, max_col=12):
        row_count += 1
        for i, cell in enumerate(row):
            if 'J' in cell.coordinate: # j column

                if isinstance(cell.value, int):
                    logging.info(str(cell.value) + '\t' + '\t\tno match! It is a number.')
                    continue
                if cell.value is None:
                    logging.info(str('Cell in column J, row {} is empty.'.format(row_count)))
                    continue # blank cell

                if isinstance(cell.value, float):
                    logging.info(str(cell.value) + '\t' + '\t\tno match! It is a float number.')
                    continue

                if isinstance(cell.value, datetime.datetime):
                    logging.info(str(cell.value) + '\t' + '\t\tno match! It is a datetime object.')
                    continue

                BLOCK_NUMBER = cell.value
                print(row_count, BLOCK_NUMBER)

                results = find_location(cell.value, BLOCK_NUMBER, FIELD_IDs)
                if results[0]:
                    continue
                else:
                    BLOCK_NUMBER = results[1]

                print('Removing any trailing letters ...')
                results = remove_trailing_letters(cell.value, BLOCK_NUMBER, FIELD_IDs)
                if results[0]:
                    continue
                else:
                    BLOCK_NUMBER = results[1]

                print('Adding in extra initials if possible ...')
                results = add_initials(cell.value, BLOCK_NUMBER, row[i+1].value, FIELD_IDs)
                if results[0]:
                    continue
                else:
                    BLOCK_NUMBER = results[1]

                print('Adding 0 to the end of the initials')
                results = add_prefix_number(cell.value, BLOCK_NUMBER, FIELD_IDs)
                if results[0]:
                    continue
                else:
                    BLOCK_NUMBER = results[1]

                not_found(cell.value, BLOCK_NUMBER)
                # input('Press enter to continue ...')

                # if not find_location(BLOCK_NUMBER, FIELD_IDs):
                #     print('\nDid not find {}'.format(BLOCK_NUMBER))
                #     input('Press enter to continue ...')
                #
                #     print('Removing any trailing letters ...')
                #     BLOCK_NUMBER = remove_trailing_letters(BLOCK_NUMBER)
                #     print('result: {}'.format(BLOCK_NUMBER))
                #
                #     if not find_location(BLOCK_NUMBER, FIELD_IDs):
                #         print('Adding in extra initials if possible ...')
                #         BLOCK_NUMBER = add_initials(BLOCK_NUMBER, row[i+1])
                #         print('result: {}'.format(BLOCK_NUMBER))
                #
                #         if not find_location(BLOCK_NUMBER, FIELD_IDs):
                #             print('Out of options')
                #             logging.info(cell.value + ' ' + BLOCK_NUMBER + '\tno match!')
                #             input('Press enter to continue ...')
                #             continue
                #
                # logging.info(cell.value + '\t' + BLOCK_NUMBER + '\tfound!')
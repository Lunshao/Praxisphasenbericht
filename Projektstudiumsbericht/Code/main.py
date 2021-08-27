try:
    from . import pandas as pd
except:
    import pandas as pd
import re
import sys


class ExcelEntry:
    """
    This class provides objects for the table entries inside of the FlashMap
    """
    def __init__(self, address, bank, collection, container, flags, img_id, img_size, img_type):
        self.address = address
        self.bank = bank
        self.collection = collection
        self.container = container
        self.flags = flags
        self.img_id = img_id
        self.img_size = img_size
        self.img_type = img_type
    

class ExcelHistory:
    """
    This class provides objects for the history entries inside the FlashMap
    """
    def __init__(self, author, date, description, version):
        self.author = author
        self.date = date
        self.description = description
        self.version = version


class ExcelContainer:
    """
    This class provides objects for the entries in Container-Sheet
    """
    def __init__(self, id, storage, number, hexid, flags):
        self.id = id
        self.storage = storage
        self.number = number
        self.hexid = hexid
        self.flags = flags
        self.size = ''

    def set_size(self, x):
        self.size = x


class ExcelPartition:
    """
    This class provides objects fot the entries in Partition-Sheer
    """
    def __init__(self, name, nr, blocks, range, type):
        self.name = name
        self.nr = nr
        self.blocks = blocks
        self.range = range
        self.type = type


def reader_excel(filename):
    """
    Reads the content from Excel-File into a Dictionary
    :param filename: Path to Excel-File
    :return: Dictionary
    """
    try:
        dict_excel = {}

        # <editor-fold desc="Overview">
        # Reads the Overview-Sheet
        df_overview = pd.read_excel(filename, 'Overview')
        bool_date = False

        for i in range(0, df_overview.shape[0] - 1):

            entry = str(df_overview.iloc[i][0])

            if entry == 'Variant':
                dict_excel['Variant'] = df_overview.iloc[i][1]

            elif entry == 'Date':
                bool_date = True
                dict_excel['History'] = []

            elif bool_date and entry != 'nan':
                dict_excel['History'].append(ExcelHistory(str(df_overview.iloc[i][1]), str(df_overview.iloc[i][0]).replace(' 00:00:00', ''),
                                                          str(df_overview.iloc[i][2]), str(df_overview.iloc[i][3])))

            elif bool_date and entry == 'nan:':
                break

        del df_overview
        del bool_date
        # </editor-fold>

        # <editor-fold desc="Collections">
        # Reads the Collections-Sheet
        df_collection = pd.read_excel(filename, 'Collection')
        bool_collection = False

        for i in range(0, df_collection.shape[0]):

            entry = str(df_collection.iloc[i][0])

            if entry == 'Collection':
                bool_collection = True
                dict_excel['Collections'] = []

            elif bool_collection and entry != 'nan':
                dict_excel['Collections'].append('COLLECTION_' + str(df_collection.iloc[i][0]) + ' ' + str(df_collection.iloc[i][2]))

            elif bool_collection and entry == 'nan':
                break

        del df_collection
        del bool_collection
        # </editor-fold>

        # <editor-fold desc="Container">
        # Reads the Container-Sheet
        df_container = pd.read_excel(filename, 'Container')
        bool_container = False

        list_containers = []

        for i in range(0, df_container.shape[0]):

            entry = str(df_container.iloc[i][1])

            if entry == 'ContainerID':

                bool_container = True

            elif bool_container and entry != 'nan':

                flags = df_container.iloc[i][5].split('\r\n')
                obj_container = ExcelContainer(df_container.iloc[i][1], df_container.iloc[i][2], df_container.iloc[i][3],
                                               df_container.iloc[i][4], flags)
                list_containers.append(obj_container)

        dict_excel['ContainerTable'] = list_containers

        del df_container
        del flags
        del list_containers
        del obj_container
        del bool_container
        # </editor-fold>

        # <editor-fold desc="Partition">
        df_partition = pd.read_excel(filename, 'Partition')
        dict_excel['Partitions'] = []

        for i in range(0, df_partition.shape[0]):

            entry = df_partition.iloc[i][0]

            if entry != '(Reserved for GPT Table)':
                partition = ExcelPartition(df_partition.iloc[i][0], df_partition.iloc[i][1], df_partition.iloc[i][3],
                                           df_partition.iloc[i][6], df_partition.iloc[i][7])
                dict_excel['Partitions'].append(partition)

            elif entry == 'nan':
                break

        del df_partition
        del partition
        del entry
        # </editor-fold>

        # <editor-fold desc="FlashMap">
        # Reads the FlashMap-Sheet
        df_flashmap = pd.read_excel(filename, 'FlashMap')
        bool_container = False
        bool_container_group = False
        bool_size = False

        dict_excel['Containers'] = []
        list_container_areas = []

        count_container = 0

        for i in range(0, df_flashmap.shape[0] - 1):

            entry = str(df_flashmap.iloc[i][3])

            # Checks if the entry is a Container Group name
            if not bool_container_group and 'Container Group' in entry:

                string_container_group = entry
                bool_container_group = True

                dict_excel[string_container_group] = []

            # Checks if the entry is the Container header of table
            elif not bool_container and 'Container' in entry:

                bool_container = True

            # Checks if an entry under the Container column is read
            elif bool_container and bool_container_group and entry != 'nan':

                if str(df_flashmap.iloc[i][5]) != 'FREE':
                    list_flags = str(df_flashmap.iloc[i][14]).split('\r\n')

                    dict_excel[string_container_group].append(ExcelEntry(df_flashmap.iloc[i][4], df_flashmap.iloc[i][12],
                                                                         df_flashmap.iloc[i][11], df_flashmap.iloc[i][3],
                                                                         list_flags, df_flashmap.iloc[i][5],
                                                                         df_flashmap.iloc[i][7], df_flashmap.iloc[i][9]))

                if entry not in dict_excel['Containers']:
                    dict_excel['Containers'].append(str(df_flashmap.iloc[i][3]))

                if 'FLI' not in str(df_flashmap.iloc[i][9]) and 'FSN' not in str(df_flashmap.iloc[i][9]):
                    bool_size = True

            # Checks if a Container Group table has ended
            elif bool_container and bool_container_group and entry == 'nan':

                if not bool_size:

                    container_size = 0
                else:

                    container_size = str(df_flashmap.iloc[i][8])
                dict_excel['ContainerTable'][count_container].set_size(container_size)

                count_container += 1
                bool_container_group = False
                bool_container = False
                bool_size = False
                list_container_areas.append(string_container_group)

        dict_excel['Container Groups'] = list_container_areas
        # </editor-fold>

        return dict_excel
    except BaseException as e:

        raise e


def reader_template(filename):
    """
    Reads a textfile as input and converts it into an dictionary
    :param filename: Path to file
    :return: Dictionary
    """
    try:

        with open(filename, 'r') as f:
            data_store = f.readlines()

        dict_data = {}
        count = 0

        for entry in data_store:
            dict_data[count] = entry
            count += 1

        return dict_data
    except BaseException as e:

        raise e


def string_builder(obj):
    """
    This method builds the string which is later written into the ImageTable
    :param obj: Information of corresponding Excel-Entry
    :return: String to write to file
    """
    try:
        # ContainerID
        container_id = '.m_uiContainerId = ' + obj.container + ',\t'

        # Address
        fill_up_digits = '0x'
        for i in range(0, 16 - len(obj.address)):

            fill_up_digits += '0'
        address = '.m_uiAddress = ' + fill_up_digits + obj.address + ',\t'

        # ImageID
        image_id = '.m_uiImageId = ' + obj.img_id + ',\t'

        # Collection
        if obj.collection != '-':

            collection = '.m_uiCollection = COLLECTION_' + obj.collection + ',\t'
        else:

            collection = '.m_uiCollection = IIO_FT_NO_COLLECTION,\t'

        # Bank
        if '-' not in str(obj.bank):

            bank = '.m_uiBank = ' + str(obj.bank) + ',\t'

        else:

            bank = '.m_uiBank = IIO_FT_NO_BANK,\t'

        # ImageType
        if 'FSN' in obj.img_type:

            image_type = '.m_eImageType = IIO_IMGTYPE_NBX,\t'

        else:

            image_type = '.m_eImageType = IIO_IMGTYPE_' + obj.img_type + ',\t'

        # ImageSize
        # Max length: 25 = 3x 4 digits, 3x UL, 5x Space, 2x *
        if int(obj.img_size) > 1024:

            size = int(int(obj.img_size)/1024)

            whitespaces = ''

            for i in range(0, 4 - len(str(size))):

                whitespaces += ' '

            image_size = str(size) + 'UL * 1024UL * 1024UL,\t'

        else:

            whitespaces = ''

            for i in range(0, 13 - len(str(obj.img_size))):

                whitespaces += ' '

            image_size = str(obj.img_size) + 'UL * 1024UL,\t'

        image_size = '.m_uiImageSize = ' + whitespaces + image_size

        # Flags
        flags = '.m_uiFlags = IIO_FT_IMAGEFLAG_' + obj.flags[0]
        del obj.flags[0]

        for flag in obj.flags:

            flag = 'IIO_FT_IMAGEFLAG_' + flag
            flags += ' | ' + flag
        flags += '},\n'

        return '\t\t\t{ ' + container_id + address + image_id + collection + bank + image_type + image_size + flags
    except BaseException as e:

        raise e


def write_history(c_file, dict_excel):
    """
    Replaces the "-- [HISTORY] --" line in template
    :param c_file: TextIOWrapper to write to file
    :param dict_excel: Dict which contains all relevant information from excel-sheet
    :return: -
    """
    c_file.write('  Date\t\t\tAuthor\t\tVersion\t\tDescription\n')

    for obj in reversed(dict_excel['History']):

        if '\n' in obj.description or len(obj.description) > 60:

            if '\n' in obj.description:

                list_entries = obj.description.split('\n')

                description = list_entries[0] + '\n'
                del list_entries[0]

                for string in list_entries:
                    description += '\t\t\t\t\t\t\t\t\t\t' + string + '\n'

                c_file.write('  ' + obj.date + '\t' + obj.author + '\t' + obj.version + '\t\t' + description)

            elif len(obj.description) > 60:

                mid = int(len(obj.description) / 2)
                count = mid - 10
                split_point = 0

                for char in obj.description[mid - 10:mid + 10:1]:

                    if char == ' ' and count < mid:

                        split_point = count

                    elif char == ' ' and count > mid:

                        split_point = count
                        break
                    count += 1
                front, back = obj.description[0:split_point], obj.description[split_point:len(obj.description)]
                c_file.write('  ' + obj.date + '\t' + obj.author + '\t' + obj.version + '\t\t' + front
                             + '\n' + '\t\t\t\t\t\t\t\t\t\t' + back.strip(' ') + '\n')

        else:

            c_file.write('  ' + obj.date + '\t' + obj.author + '\t' + obj.version + '\t\t' + obj.description + '\n')


def write_version(c_file, string_version):
    """
    Replaces the "-- [VERSION] --" line in template
    :param c_file:
    :param string_version:
    :return:
    """
    string_major = re.match('^(\d.\d.\d).(\d)$', string_version).group(1)
    c_file.write('#define IIO_DFT_VERSION "' + string_major + '"\n')


def write_version_minor(c_file, string_version):
    """
    Replaces the "-- [MINORVERSION] --" line in template
    :param c_file:
    :param string_version:
    :return:
    """
    string_minor = re.match('^(\d.\d.\d).(\d)$', string_version).group(2)
    c_file.write('#define IIO_DFT_MINOR_VERSION "' + string_minor + '"\n')


def write_collections(c_file, dict_excel):
    """
    Replaces the "-- [COLLECTIONS] --" line in template
    :param c_file:
    :param dict_excel:
    :return:
    """
    for string in dict_excel['Collections']:
        c_file.write('#define ' + string + '\n')


def write_containers(c_file, dict_excel, dict_template, key):
    """
    Replaces the "-- [CONTAINERS] --" line in template
    :param c_file:
    :param dict_excel:
    :param dict_template:
    :param key:
    :return:
    """
    for container in dict_excel['ContainerTable']:

        c_file.write('#define CONTAINER_' + container.id + ' ' + container.hexid + '\n')


def write_variantversion(c_file, dict_excel, value):
    """
    Replaces the "-- [VARIANTVERSION] --" part in line with variant
    :param c_file:
    :param dict_excel:
    :return:
    """
    version = dict_excel['Variant'] + ' v'
    c_file.write(value.replace('-- [VARIANTVERSION] --', version))


def write_containertable(c_file, dict_excel):
    """
    Replaces the "-- [CONTAINERTABLE] --" line in template
    :param c_file:
    :param dict_excel:
    :return:
    """
    for container in dict_excel['ContainerTable']:

        # ContainerID
        id = '.m_uiContainerId = CONTAINER_' + container.id + ',\t'
        # Storage
        storage = '.m_eStorage = ' + container.storage + ',\t'
        # Number
        # number = '.m_uiNumber = ' + container.number + ',\t'
        number = '.m_uiNumber = ' + str(container.number) + ',\t'
        # ContainerSize
        if container.size == 0:
            size = '.m_uiContainerSize = \t\t\t\t\t   ' + str(container.size) + ',\t'
        else:
            size = '.m_uiContainerSize = ' + str(container.size) + 'UL * 1024UL * 1024UL,\t'
        # Flags
        flags = '.m_uiFlags = '
        for flag in container.flags:

            flags += flag + ' | '

        flags = flags.strip(' | ')

        c_file.write('\t\t\t{ ' + id + storage + number + size + flags + ' },\n')


def write_tableentries(c_file, dict_excel):
    """
    Replaces the "-- [TABLEENTRIES] --" line in template
    :param c_file:
    :param dict_excel:
    :return:
    """
    string_table_entries = '\t\t.m_uiNofDefaultImagetableEntries = '

    for container in dict_excel['Container Groups']:
        string_table_entries += str(len(dict_excel[container])) + ' + '

    string_table_entries = string_table_entries.strip(' + ')
    string_table_entries += ',\n'
    c_file.write(string_table_entries)


def write_imagetable(c_file, dict_excel):
    """
    Replaces the "-- [IMAGETABLE] --" line in template
    :param c_file:
    :param dict_excel:
    :return:
    """
    for container in dict_excel['Container Groups']:

        c_file.write('\t\t\t/* Images on ' + container + ' */\n')

        if len(dict_excel[container]) != 0:

            for obj in dict_excel[container]:
                c_file.write(string_builder(obj))

            c_file.write('\n')

        else:

            c_file.write('\t\t\t/* Free area */\n\n')

    c_file.write('\t\t}\n\t},\n')


def handler(excel_filename, template_filename, template_shell):
    """
    Handles the writing to the C-File and replacement of placeholder lines in template
    :param excel_filename:
    :param template_filename:
    :param template_shell:
    :return:
    """
    try:
        dict_excel = reader_excel(excel_filename)
        dict_template = reader_template(template_filename)
        del excel_filename
        del template_filename

        string_version = dict_excel['History'][len(dict_excel['History']) - 1].version

        string_c_filename = 'files\\iio_cfg_iip_kilimanjaro_' + string_version + '.c'
        c_file = open(string_c_filename, 'w')
        del string_c_filename

        for key in dict_template:
            value = dict_template[key]

            if '-- [HISTORY] --' in value:
                write_history(c_file, dict_excel)

            elif '-- [VERSION] --' in value:
                write_version(c_file, string_version)

            elif '-- [MINORVERSION] --' in value:
                write_version_minor(c_file, string_version)

            elif '-- [COLLECTIONS] --' in value:
                write_collections(c_file, dict_excel)

            elif '-- [CONTAINERS] --' in value:
                write_containers(c_file, dict_excel, dict_template, key)

            elif '-- [VARIANTVERSION] --' in value:
                write_variantversion(c_file, dict_excel, value)

            elif '-- [CONTAINERTABLE] --' in value:
                write_containertable(c_file, dict_excel)

            elif '-- [TABLEENTRIES] --' in value:
                write_tableentries(c_file, dict_excel)

            elif '-- [IMAGETABLE] --' in value:
                write_imagetable(c_file, dict_excel)

            else:
                c_file.write(value)

        handler_shell(template_shell, dict_excel, string_version)

    except BaseException as e:

        raise e


def handler_shell(shell_filename, dict_excel, version):
    """
    Reads the shell template and creates the new shell-script
    :param shell_filename: Path to shell-template
    :param dict_excel: Dictionary containing information from excel-sheet
    :param version: String of current build version
    :return: -
    """
    dict_shell_template = reader_template(shell_filename)
    shell_file = open('files\\iio-mmc-part-layout.sh', 'w')

    for key in dict_shell_template:
        value = dict_shell_template[key]

        if '-- [VERSION] --' in value:

            replacement = value.replace('-- [VERSION] --', dict_excel['Variant'] + ' v' + version)
            shell_file.write(replacement)

        elif '-- [PARTITIONS] --' in value:

            for partition in dict_excel['Partitions']:

                part = 'part' + str(partition.nr) + '_'

                shell_file.write('# Size in 512B blocks: ' + str(partition.blocks) + ' blocks\n')
                shell_file.write(part + 'name="' + partition.name + '"\n')
                shell_file.write(part + 'sectors="' + partition.range + '"\n')
                shell_file.write(part + 'fs="' + partition.type + '"\n\n')

            shell_file.write('nof_partitions=' + str(dict_excel['Partitions'].__len__()) + '\n')

        else:
            shell_file.write(value)


def main():
    handler(sys.argv[1], sys.argv[2], sys.argv[3])


if __name__ == "__main__":
    main()

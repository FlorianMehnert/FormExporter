import pathlib

import pandas as pd


class Pair:
    def __init__(self, one, two):
        self.one = one
        self.two = two


def fill_second_parameter(parameters: Pair, ls, index):
    bold_lines = get_bold_line_numbers(ls)
    if parameters.one in ls[index]:
        ul = find_next_number(bold_lines, index)
        parameters.two = construct_content_of_lines(ls, ul, index)


def create_df_row(p: Pair):
    return p.two,


class Form:
    def __init__(self):
        self.vorname = Pair("Vorname", None)  # titel, lokaler content
        self.nachname = Pair("Nachname", None)
        self.bestelle_als = Pair("Ich bestelle als", None)
        self.einrichtung = Pair("Einrichtung (optional)", None)
        self.strasse = Pair("Straße", None)
        self.hausnummer = Pair("Hausnummer", None)
        self.plz = Pair("PLZ", None)
        self.stadt = Pair("Stadt", None)
        self.email = Pair("E-Mail", None)
        self.telefon = Pair("Telefon (optional)", None)
        self.erfahren = Pair("Wie haben Sie von dem Projekt erfahren? (optional)", None)
        self.mitteilen = Pair("Möchten Sie uns noch etwas mitteilen? (optional)", None)

    def get_dataframe(self):
        return [create_df_row(self.vorname),
                create_df_row(self.nachname),
                create_df_row(self.bestelle_als),
                create_df_row(self.einrichtung),
                create_df_row(self.strasse),
                create_df_row(self.hausnummer),
                create_df_row(self.plz),
                create_df_row(self.stadt),
                create_df_row(self.email),
                create_df_row(self.telefon),
                create_df_row(self.erfahren),
                create_df_row(self.mitteilen)]


def is_bold(l: str):
    l = l.strip()
    return l.startswith("**") and l.endswith("**")


def get_bold_line_numbers(lines: list):
    """
    this function finds all lines that contain a bold word marking a heading
    :return:
    """
    numbers = []
    for i, line in enumerate(lines):
        if is_bold(line):
            numbers.append(i)
    numbers.append(len(lines) + 1)
    return numbers


def find_next_number(number_list, number):
    for current_number in number_list:
        if number < current_number:
            return current_number


def construct_content_of_lines(lines: list[str], until, starting):
    content = ""
    for i, line in enumerate(lines):
        if until > i > starting:
            if i == until - 1:
                content += line.replace("\n", "")
            else:
                content += line
    return content


if __name__ == '__main__':
    form = Form()
    with open("foo.md", "r") as f:
        current_topic = ""
        lines = f.readlines()
        for i, line in enumerate(lines):
            extra_parameters = (lines, i)
            fill_second_parameter(form.vorname, *extra_parameters)
            fill_second_parameter(form.nachname, *extra_parameters)
            fill_second_parameter(form.bestelle_als, *extra_parameters)
            fill_second_parameter(form.einrichtung, *extra_parameters)
            fill_second_parameter(form.telefon, *extra_parameters)
            fill_second_parameter(form.email, *extra_parameters)
            fill_second_parameter(form.stadt, *extra_parameters)
            fill_second_parameter(form.plz, *extra_parameters)
            fill_second_parameter(form.erfahren, *extra_parameters)
            fill_second_parameter(form.mitteilen, *extra_parameters)
            fill_second_parameter(form.strasse, *extra_parameters)
            fill_second_parameter(form.hausnummer, *extra_parameters)
    excel_file = pathlib.Path("excel_file.xlsx")
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df = pd.DataFrame.from_records(form.get_dataframe()).transpose()
        # df.to_excel(excel_writer=writer, index=False, sheet_name="Sheet1", header=False)
        print(df.to_excel(writer, sheet_name="Sheet1", startrow=writer.sheets["Sheet1"].max_row, index=False, header=False))

        writer._save()

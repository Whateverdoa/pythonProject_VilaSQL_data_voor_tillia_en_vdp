# This is a sample Python script.
from pathlib import Path
import pandas as pd
import openpyxl
from connect_mssql_vila import pd30X30_inner

"""data komt als file uit VILA sql"""

filename = "uit_SQL-TEST-excel/leesfile.xlsx"
pad = Path(filename)


def file_to_generator(file_in: Path) -> pd.DataFrame:
    if Path(file_in).suffix == ".csv":
        file_to_generate_on = pd.read_csv(file_in, ";")

    elif Path(file_in).suffix == ".xlsx":
        print(Path(file_in).suffix)
        file_to_generate_on = pd.read_excel(file_in, engine="openpyxl")

    elif Path(file_in).suffix == ".xls":
        print(Path(file_in).suffix)
        file_to_generate_on = pd.read_excel(file_in)

    # file_to_generate_on.replace(['nan', 'None'], '')
    return file_to_generate_on


# bewerk eerst nieuwe_df
final_dataframe = file_to_generator(pad).replace(["nan", "NaN", "None", "."], "")

print(*final_dataframe.columns)

# generator = function.file_to_generator(pad).itertuples(index=True)
generator = final_dataframe.itertuples(
    index=True,
)


def rol_wikkeling_matches(rwid: int) -> str:
    match rwid:
        case 1:
            return "geen Rolwikkeling"
        case 2:
            return "Rolwikkeling 1"
        case 3:
            return "Rolwikkeling 2"
        case 4:
            return "Rolwikkeling 3"
        case 5:
            return "Rolwikkeling 4"
        case 6:
            return "Rolwikkeling 5"
        case 7:
            return "Rolwikkeling 6"
        case 8:
            return "Rolwikkeling 7"
        case 9:
            return "Rolwikkeling 8"
        case 10:
            return "Rolwikkeling diversen"


# Nota Bena aantal = aantal_per_rol
kolommen = [
    "ordernummer",
    "beeld",
    "SKU",
    "aantal",
    "aantal_rollen",
    "rolwikkeling",
    "kernID",
    "verstuurdatum",
    "klantnaam",
    "vorm",
    "CuttingDie",
    "LabelWidth",
    "LabelHeight",
    "batchcover"
]


def rows_to_line_in_df(rows_from_itertuple: iter) -> pd.DataFrame:
    regel1 = pd.DataFrame(
        [
            (
                f"{rows_from_itertuple.ordernummer}",
                f"{rows_from_itertuple.itemnummer}.pdf",
                f"{rows_from_itertuple.ClientOrderNo}",
                f"{rows_from_itertuple.aantal_per_rol}",
                f"{rows_from_itertuple.aantal_rollen}",
                f"{rol_wikkeling_matches(int(rows_from_itertuple.rolwikkeling))}",
                f"{rows_from_itertuple.kernID}",
                f"{rows_from_itertuple.verstuurdatum}",
                f"{rows_from_itertuple.klantnaam}",
                f"{rows_from_itertuple.vorm}",
                f"{rows_from_itertuple.CuttingDie}",
                f"{rows_from_itertuple.LabelWidth}",
                f"{rows_from_itertuple.LabelHeight}",
                "False"
            )
            for i in range(rows_from_itertuple.aantal_rollen)
        ],
        columns=kolommen,
    )
    regel2 = pd.DataFrame(
        [
            (
                f"{rows_from_itertuple.ordernummer}",
                "leeg.pdf",
                f"{rows_from_itertuple.ClientOrderNo}",
                f"{rows_from_itertuple.aantal_per_rol}",
                "1",
                f"{rol_wikkeling_matches(int(rows_from_itertuple.rolwikkeling))}",
                f"{rows_from_itertuple.kernID}",
                f"{rows_from_itertuple.verstuurdatum}",
                f"{rows_from_itertuple.klantnaam}",
                f"{rows_from_itertuple.vorm}",
                f"{rows_from_itertuple.CuttingDie}",
                f"{rows_from_itertuple.LabelWidth}",
                f"{rows_from_itertuple.LabelHeight}",
                "False"
            )
        ],
        columns=kolommen,
    )

    regels = pd.concat([regel2, regel1, regel2], axis=0)

    return regel1


def adding_lead_in_lead_out(rows_from_itertuple: iter) -> pd.DataFrame:
    regel1 = pd.DataFrame(
        [
            (
                f"{rows_from_itertuple.ordernummer}",
                f"{rows_from_itertuple.beeld}",
                f"{rows_from_itertuple.SKU}",
                f"{rows_from_itertuple.aantal}",
                "1",
                f"{rows_from_itertuple.rolwikkeling}",
                f"{rows_from_itertuple.kernID}",
                f"{rows_from_itertuple.verstuurdatum}",
                f"{rows_from_itertuple.klantnaam}",
                f"{rows_from_itertuple.vorm}",
                f"{rows_from_itertuple.CuttingDie}",
                f"{rows_from_itertuple.LabelWidth}",
                f"{rows_from_itertuple.LabelHeight}",
                "False"
            )
        ],
        columns=kolommen,
    )

    regel2 = pd.DataFrame(
        [
            (
                f"{rows_from_itertuple.ordernummer}",
                "leeg.pdf",
                f"{rows_from_itertuple.SKU}",
                "1",
                "1",
                f"{rows_from_itertuple.rolwikkeling}",
                f"{rows_from_itertuple.kernID}",
                f"{rows_from_itertuple.verstuurdatum}",
                f"{rows_from_itertuple.klantnaam}",
                f"{rows_from_itertuple.vorm}",
                f"{rows_from_itertuple.CuttingDie}",
                f"{rows_from_itertuple.LabelWidth}",
                f"{rows_from_itertuple.LabelHeight}",
                "True"
            )
        ],
        columns=kolommen,
    )

    regel3 = pd.DataFrame(
        [
            (
                f"{rows_from_itertuple.ordernummer}",
                "stans.pdf",
                f"{rows_from_itertuple.SKU}",
                "1",
                "1",
                f"{rows_from_itertuple.rolwikkeling}",
                f"{rows_from_itertuple.kernID}",
                f"{rows_from_itertuple.verstuurdatum}",
                f"{rows_from_itertuple.klantnaam}",
                f"{rows_from_itertuple.vorm}",
                f"{rows_from_itertuple.CuttingDie}",
                f"{rows_from_itertuple.LabelWidth}",
                f"{rows_from_itertuple.LabelHeight}",
                "False"
            )
        ],
        columns=kolommen,
    )

    number_of_columns = len(kolommen)
    filling_for_true_leeg_stans = ["" for i in range(number_of_columns)]
    filling_for_true_leeg_stans[1] = "stans.pdf"  # or "stans"

    filling_for_true_leeg_stans[3] = 1
    blanco_regel = pd.DataFrame([filling_for_true_leeg_stans], columns=kolommen)

    # regels = pd.concat([blanco_regel, regel2, regel1, regel2, blanco_regel], axis=0)
    regels = pd.concat([regel3, regel2, regel1 ,regel2, regel3], axis=0)
    return regels


# SalesOrderNo,ShapeId


# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    # verwerkte file(.xlsx) in zou moeten werken in df naar csv
    verwerkte_file_in = pd.concat([rows_to_line_in_df(rows) for rows in generator])

    df_met_lead_in_out = pd.concat(
        [
            adding_lead_in_lead_out(rows)
            for rows in verwerkte_file_in.itertuples(index=False)
        ]
    )

    df_met_lead_in_out.to_excel("uit_SQL-TEST-excel/vdp_or_tilia_file.xlsx", index=False)
    print(df_met_lead_in_out.head(4))
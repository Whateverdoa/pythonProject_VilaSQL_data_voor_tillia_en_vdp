from pathlib import Path
from dataclasses import dataclass

klantnamen= ["HELLOPRINT.B.V", "DRUKWERKDEAL.NL", "PRINT.COM"]
PRODUCTS_path = Path("M:\ESKO\Products")



@dataclass
class Pad_vila_itemnummers:

    klantnamen: list


def pad_naar_item_nummer_folder_maker(voor_itemnummer_uit_lijst):
    """Builds a path filename from the vila item number for the
    Helloprint, Drukwerkdeal and print.com reseller clients.
    # hp_id = "806321"
    # dwd_id = "569621"
    # pdc_id = "935321
    """
    voor_itemnummer_uit_lijst = str(voor_itemnummer_uit_lijst)
    itemnummer_pdf = voor_itemnummer_uit_lijst + ".pdf"

    def foldername_based_on_itemnummer(itemnummer):
        itemnummer = str(itemnummer)
        folderbase = itemnummer[4:9]
        foldername = folderbase + "000-" + folderbase + "999"
        return foldername

    def klantnummer_uit(itemnummer):
        itemnummer = str(itemnummer)
        klantnummer = itemnummer[0:6]
        return klantnummer

    klant_naam_paden = [(PRODUCTS_path.joinpath(name[0], name)) for name in klantnamen]

    helloprint_base_path = klant_naam_paden[0]
    drukwerkdeal_base_path = klant_naam_paden[1]
    print_dot_com_base_path = klant_naam_paden[2]

    foldername = foldername_based_on_itemnummer(voor_itemnummer_uit_lijst)

    match klantnummer_uit(voor_itemnummer_uit_lijst):
        case "806321":
            return Path(helloprint_base_path.joinpath(foldername, voor_itemnummer_uit_lijst, itemnummer_pdf))

        case "569621":
            return Path(drukwerkdeal_base_path.joinpath(foldername, voor_itemnummer_uit_lijst, itemnummer_pdf))

        case "935321":
            return Path(print_dot_com_base_path.joinpath(foldername, voor_itemnummer_uit_lijst, itemnummer_pdf))
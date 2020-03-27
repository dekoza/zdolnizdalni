import os
from pathlib import Path

import pandas as pd

names = [
    "LP",
    "RSPO",
    "pol_woj",
    "pol_pow",
    "pol_gm",
    "miejscowosc",
    "klasa_wielk",
    "typ",
    "zlozonosc",
    "nazwa",
    "patron",
    "ulica",
    "numer_domu",
    "kod_poczt",
    "poczta",
    "telefon",
    "fax",
    "www",
    "publicznosc",
    "kat_uczniow",
    "spec_szkoly",
    "zwiaz_org",
    "op_ko",
    "op_nazwa",
    "op_woj",
    "op_pow",
    "op_gm",
    "typ_gminy",
    "or_kod",
    "or_nazwa",
    "or_woj",
    "or_pow",
    "or_gm",
    "regon",
    "uczniowie",
    "female",
    "w_przedszk",
    "w_innych",
    "oddzialow",
    "naucz_full",
    "naucz_sp",
    "naucz_etat",
]
keep = [
    "LP",
    "miejscowosc",
    "nazwa",
    "patron",
    "ulica",
    "numer_domu",
    "telefon",
    "fax",
    "www",
]
to_drop = set(names) - set(keep)


def read_file(path):
    return pd.read_excel(path, skiprows=2, header=6, index_col=0, names=names,)


def drop_unneeded(dataframe):
    dataframe = dataframe.loc[lambda x: x.publicznosc == 1].loc[
        lambda x: x.kat_uczniow == 1
    ]

    df1 = dataframe.loc[lambda x: x.typ == 3]
    df2 = dataframe.loc[lambda x: x.typ == 14]

    return df1.merge(df2, how="outer").drop(columns=to_drop)


def main():
    inpath = Path("./input")
    outpath = Path("./output")
    incoming = os.listdir("./input/")
    with pd.ExcelWriter(outpath / "result.xls") as writer:
        for file in (f for f in incoming if f.endswith(".xls")):
            name, _ = file.split(".")
            df = drop_unneeded(read_file(inpath / file))
            df.to_excel(writer, sheet_name=name)


if __name__ == "__main__":
    main()

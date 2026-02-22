import pandas as pd
import re
from datetime import datetime

def inverse_generate_ecdv(ecdv_string: str) -> pd.DataFrame:

    if not isinstance(ecdv_string, str):
        raise TypeError("ECDV input must be a string.")

    ecdv_string = ecdv_string.strip()

    if not ecdv_string:
        raise ValueError("Empty ECDV string.")

    if ecdv_string == "No combinations for this product line":
        raise ValueError("Cannot inverse: No combinations case.")

    if not ecdv_string.endswith("*"):
        raise ValueError("Invalid ECDV format (missing '*').")

    ecdv_string = ecdv_string[:-1]

    match = re.match(r'^([^.]+)\.([A-Za-z0-9]+)(.*)$', ecdv_string)

    if not match:
        raise ValueError("Invalid ECDV structure.")

    CM = match.group(1)
    Family = match.group(2)
    remainder = match.group(3)

    if remainder.startswith("."):
        remainder = remainder[1:]

    if "<" in remainder:
        common_str, body = remainder.split("<", 1)
        common_parts = [p for p in common_str.split(".") if p]
    else:
        common_parts = []
        body = remainder

    combinations = body.split("/") if body else []
    parsed_rows = []

    for combo in combinations:

        combo = combo.strip()
        if not combo:
            continue

        row_dict = {}

        tokens = re.findall(
            r"\([A-Z0-9]+[A-Z0-9]{2}\)|[A-Z0-9]+[A-Z0-9]{2}",
            combo
        )

        for token in tokens:

            is_exception = False

            if token.startswith("("):
                is_exception = True
                token = token[1:-1]

            col = token[:-2]
            val = token[-2:]

            if is_exception:
                val = f"!{val}"

            if col in row_dict:

                existing = row_dict[col]

                if not isinstance(existing, list):
                    existing = [existing]

                if any(not v.startswith("!") for v in existing):
                    raise ValueError(
                        f"Invalid ECDV: mixed inclusion/exclusion for column {col}"
                    )

                if not val.startswith("!"):
                    raise ValueError(
                        f"Invalid ECDV: mixed inclusion/exclusion for column {col}"
                    )

                existing.append(val)
                row_dict[col] = existing

            else:
                row_dict[col] = val

        parsed_rows.append(row_dict)

    if not parsed_rows:
        raise ValueError("No valid combinations parsed.")

    for row in parsed_rows:
        for part in common_parts:
            col = part[:-2]
            val = part[-2:]
            row[col] = val

    all_columns = sorted({col for row in parsed_rows for col in row.keys()})

    final_rows = []
    for row in parsed_rows:
        formatted = {}
        for col in all_columns:
            formatted[col] = row.get(col, [])
        final_rows.append(formatted)

    return pd.DataFrame(final_rows)


def preprocess_ecdv_for_comparison(ecdv1, ecdv2):

    df1 = inverse_generate_ecdv(ecdv1)
    df2 = inverse_generate_ecdv(ecdv2)

    all_columns = sorted(set(df1.columns).union(set(df2.columns)))

    for col in all_columns:
        if col not in df1.columns:
            df1[col] = [[] for _ in range(len(df1))]

    for col in all_columns:
        if col not in df2.columns:
            df2[col] = [[] for _ in range(len(df2))]

    df1 = df1[all_columns]
    df2 = df2[all_columns]

    return df1, df2


def normalize_cell(val):

    if val == []:
        return []

    if isinstance(val, list):
        return [str(v) for v in val]

    return [str(val)]


def is_exclusion(v):
    return v.startswith("!")


def is_inclusion(v):
    return not v.startswith("!")


def rows_are_duplicate(row1, row2, columns):

    for col in columns:

        vals1 = normalize_cell(row1[col])
        vals2 = normalize_cell(row2[col])

        if not vals1 or not vals2:
            continue

        if (
            all(is_inclusion(v) for v in vals1)
            and all(is_inclusion(v) for v in vals2)
        ):
            if vals1[0] != vals2[0]:
                return False

    for col in columns:

        vals1 = normalize_cell(row1[col])
        vals2 = normalize_cell(row2[col])

        if not vals1 or not vals2:
            continue

        if any(is_exclusion(v) for v in vals1) and all(is_inclusion(v) for v in vals2):
            incl = vals2[0]
            if f"!{incl}" in vals1:
                return False

        if any(is_exclusion(v) for v in vals2) and all(is_inclusion(v) for v in vals1):
            incl = vals1[0]
            if f"!{incl}" in vals2:
                return False

    return True


def row_to_combination_string(row):

    parts = []

    for col, val in row.items():

        if val == []:
            continue

        if not isinstance(val, list):
            val = [val]

        for v in val:
            v = str(v)

            if v.startswith("!"):
                parts.append(f"({col}{v[1:]})")
            else:
                parts.append(f"{col}{v}")

    return ".".join(parts) if parts else "ALL"


def find_duplicates_one_to_many(
        new_ecdv,
        other_ecdvs,
        new_product_number=None,
        other_product_numbers=None
):

    duplicates_global = False

    for idx, ecdv in enumerate(other_ecdvs):

        if new_product_number and other_product_numbers:
            if new_product_number == other_product_numbers[idx]:
                continue

        df1, df2 = preprocess_ecdv_for_comparison(
            new_ecdv,
            ecdv
        )

        columns = df1.columns
        duplicate_pairs = []

        for _, row1 in df1.iterrows():
            for _, row2 in df2.iterrows():

                if rows_are_duplicate(row1, row2, columns):

                    combo1 = row_to_combination_string(row1)
                    combo2 = row_to_combination_string(row2)

                    duplicate_pairs.append(f"{combo1} and {combo2}")

        if duplicate_pairs:

            duplicates_global = True
            unique_pairs = list(dict.fromkeys(duplicate_pairs))

            if new_product_number and other_product_numbers:
                part1 = f"ref. {new_product_number}"
                part2 = f"ref. {other_product_numbers[idx]}"
            else:
                part1 = "part 1"
                part2 = f"part {idx+2}"

            print(
                f"\nBetween {part1} and {part2} the following combinations are forming duplicates: "
                + ", ".join(unique_pairs) + "."
            )

    if not duplicates_global:
        print("\nNo duplicates are forming with the existing parts.")

def normalize_excel_ecdv_format(ecdv: str):

    if not isinstance(ecdv, str):
        return ecdv

    ecdv = ecdv.strip()

    if not ecdv:
        return ecdv

    # Remove B0 / D / F after . / < ( )
    ecdv = re.sub(
        r'(?<=[\./<\(\)])(?:B0|D|F)(?=[A-Z0-9])',
        '',
        ecdv
    )

    return ecdv

def load_excel_master_dataframe(file_path):
    """
    Load Excel only once per day.
    Keeps original dataframe untouched.

    Fixes:
    - Mixed date formats
    - Open-ended end dates (NaT → 2999-12-31)
    """

    df_master = pd.read_excel(
        file_path,
        header=1,
        dtype=str
    )

    required_columns = [
        "05 Numero produit",
        "02 Code fonction lien vehicule",
        "ECDV",
        "Date application OEV debut",
        "Date application OEV fin"
    ]

    df_master = df_master[required_columns].copy()

    # ---------------------------
    # Convert dates (robust parse)
    # ---------------------------

    df_master["Date application OEV debut"] = pd.to_datetime(
        df_master["Date application OEV debut"],
        errors="coerce",
        dayfirst=True,
        format="mixed"
    )

    df_master["Date application OEV fin"] = pd.to_datetime(
        df_master["Date application OEV fin"],
        errors="coerce",
        dayfirst=True,
        format="mixed"
    )

    # ---------------------------------------------------
    # CRITICAL FIX:
    # Treat NaT end-dates as open-ended validity
    # ---------------------------------------------------

    OPEN_END_DATE = pd.Timestamp.max

    df_master["Date application OEV fin"] = df_master[
        "Date application OEV fin"
    ].fillna(OPEN_END_DATE)

    return df_master

def extract_filtered_excel_inputs(
    df_master,
    code_function,
    new_product_NFCdate
):

    date_value = pd.to_datetime(new_product_NFCdate)

    df_filtered = df_master.copy()

    df_filtered = df_filtered[
        df_filtered["02 Code fonction lien vehicule"] == code_function
    ]

    df_filtered = df_filtered[
        (df_filtered["Date application OEV debut"] <= date_value) &
        (df_filtered["Date application OEV fin"] >= date_value)
    ]

    other_product_numbers = []
    other_ecdvs = []

    for _, row in df_filtered.iterrows():

        product = str(row["05 Numero produit"]).strip()
        ecdv = normalize_excel_ecdv_format(row["ECDV"])

        if product and ecdv:
            other_product_numbers.append(product)
            other_ecdvs.append(ecdv)


    return other_product_numbers, other_ecdvs

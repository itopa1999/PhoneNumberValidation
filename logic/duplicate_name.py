import pandas as pd
from datetime import timedelta

def process_duplicates(file_path):
    """
    Processes the attendance Excel file to remove duplicates.
    Returns: (grouped_dataframe, original_count, correction_count)
    """
    df = pd.read_excel(file_path)
    required_cols = ["Name", "Join", "Left", "Duration"]
    
    if not all(col in df.columns for col in required_cols):
        raise ValueError(f"Excel must contain columns: {', '.join(required_cols)}")

    # ---------- Parse Join and Left with explicit format ----------
    date_format = "%m/%d/%Y, %I:%M:%S %p"
    df["Join_dt"] = pd.to_datetime(df["Join"], format=date_format, errors='coerce')
    df["Left_dt"] = pd.to_datetime(df["Left"], format=date_format, errors='coerce')

    # Parse existing Duration
    def parse_duration(dur_str):
        if isinstance(dur_str, timedelta):
            return dur_str
        try:
            parts = str(dur_str).strip().split(":")
            if len(parts) == 3:
                h, m, s = map(int, parts)
                return timedelta(hours=h, minutes=m, seconds=s)
            elif len(parts) == 2:
                m, s = map(int, parts)
                return timedelta(minutes=m, seconds=s)
            else:
                return timedelta(0)
        except:
            return timedelta(0)

    df["Duration_TD"] = df["Duration"].apply(parse_duration)

    # Validate and correct durations
    df["True_Duration_TD"] = df["Left_dt"] - df["Join_dt"]
    correction_count = 0
    
    for idx, row in df.iterrows():
        true_td = row["True_Duration_TD"]
        given_td = row["Duration_TD"]
        if pd.notna(true_td):
            if abs((true_td - given_td).total_seconds()) > 1:
                df.at[idx, "Duration_TD"] = true_td
                correction_count += 1

    # Group by Name
    grouped = df.groupby("Name", as_index=False).agg({
        "Join": "first",
        "Left": "last",
        "Duration_TD": "sum"
    })

    # Format duration as "H:MM:SS"
    def format_timedelta(td):
        total_seconds = int(td.total_seconds())
        hours, rem = divmod(total_seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        return f"{hours}:{minutes:02d}:{seconds:02d}"

    grouped["Duration"] = grouped["Duration_TD"].apply(format_timedelta)
    grouped.drop("Duration_TD", axis=1, inplace=True)
    grouped = grouped[["Name", "Join", "Left", "Duration"]]

    return grouped, len(df), correction_count
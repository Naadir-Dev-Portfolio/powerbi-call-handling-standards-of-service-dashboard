from __future__ import annotations

import json
import math
import shutil
import uuid
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "Source Data" / "raw"
CURATED_DIR = ROOT / "Curated Data"
SEMANTIC_DIR = ROOT / "Call Handling Standards of Service Dashboard.SemanticModel"
REPORT_DIR = ROOT / "Call Handling Standards of Service Dashboard.Report"
DEFINITION_DIR = SEMANTIC_DIR / "definition"
TABLES_DIR = DEFINITION_DIR / "tables"
PAGES_DIR = REPORT_DIR / "definition" / "pages"

REPORT_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"


def stable_guid(name: str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"call-handling-sos/{name}"))


def clean_text(value: object, fallback: str = "Unknown") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    return text if text else fallback


def normalise_time(value: object) -> str | None:
    if pd.isna(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{value.hour:02d}:{value.minute:02d}:00"
    text = str(value).strip()
    if not text:
        return None
    try:
        td = pd.to_timedelta(text)
    except ValueError:
        return text
    total_seconds = int(td.total_seconds())
    hour = total_seconds // 3600
    minute = (total_seconds % 3600) // 60
    return f"{hour:02d}:{minute:02d}:00"


def make_key_lookup(values: list[str]) -> dict[str, int]:
    unique_values = []
    seen = set()
    for value in values:
        cleaned = clean_text(value)
        if cleaned not in seen:
            seen.add(cleaned)
            unique_values.append(cleaned)
    return {value: idx + 1 for idx, value in enumerate(unique_values)}


def write_csv(df: pd.DataFrame, filename: str) -> None:
    CURATED_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(CURATED_DIR / filename, index=False, encoding="utf-8")


def build_curated_data() -> dict[str, object]:
    data_path = RAW_DIR / "data.xlsx"
    mapping_path = RAW_DIR / "mapping table.xlsx"

    data = pd.read_excel(data_path, sheet_name="Data")
    call_type_map = pd.read_excel(mapping_path, sheet_name="Mapping - Call Types")
    agent_team_map = pd.read_excel(mapping_path, sheet_name="Mapping - Agent Team")

    source_row_count = len(data)
    invalid_date_rows = data[data["Date"].isna()].copy()
    fact_source = data[data["Date"].notna()].copy()

    call_type_map = call_type_map.rename(
        columns={
            "Call_Type": "Call Type",
            "Team_Group": "Call Group",
        }
    )
    for column in ["Call Type", "Team", "Call Group", "Function"]:
        call_type_map[column] = call_type_map[column].map(lambda value: clean_text(value, "Unmapped"))

    fact_source["Call Type"] = fact_source["Call Type"].map(lambda value: clean_text(value, "Unmapped"))
    fact_source["Call Date"] = pd.to_datetime(fact_source["Date"]).dt.date
    fact_source["Half Hour"] = fact_source["Half Hour"].map(normalise_time)

    merged = fact_source.merge(call_type_map, on="Call Type", how="left")
    for column in ["Team", "Call Group", "Function"]:
        merged[column] = merged[column].map(lambda value: clean_text(value, "Unmapped"))

    merged["Queue"] = merged["Call Type"]
    merged["Emergency Classification"] = merged.apply(
        lambda row: "Emergency"
        if row["Team"] == "Emergency"
        or "Emergency" in row["Call Type"]
        or row["Call Type"] == "False Alarm"
        else "Non-Emergency",
        axis=1,
    )
    merged["Working Hours"] = merged["In Hours"].map(lambda value: "In Hours" if int(value) == 1 else "Out of Hours")

    call_type_lookup = make_key_lookup(call_type_map["Call Type"].tolist())
    team_lookup = make_key_lookup(sorted(merged["Team"].dropna().unique().tolist()))
    call_group_lookup = make_key_lookup(sorted(merged["Call Group"].dropna().unique().tolist()))
    queue_lookup = make_key_lookup(call_type_map["Call Type"].tolist())
    function_lookup = make_key_lookup(sorted(merged["Function"].dropna().unique().tolist()))
    emergency_lookup = {"Emergency": 1, "Non-Emergency": 2}
    working_hours_lookup = {"Out of Hours": 0, "In Hours": 1}

    merged["CallTypeKey"] = merged["Call Type"].map(call_type_lookup)
    merged["TeamKey"] = merged["Team"].map(team_lookup)
    merged["CallGroupKey"] = merged["Call Group"].map(call_group_lookup)
    merged["QueueKey"] = merged["Queue"].map(queue_lookup)
    merged["FunctionKey"] = merged["Function"].map(function_lookup)
    merged["EmergencyClassKey"] = merged["Emergency Classification"].map(emergency_lookup)
    merged["WorkingHoursKey"] = merged["Working Hours"].map(working_hours_lookup)
    merged["DateKey"] = pd.to_datetime(merged["Call Date"]).dt.strftime("%Y%m%d").astype(int)

    time_values = sorted(merged["Half Hour"].dropna().unique().tolist())
    time_lookup = {value: idx + 1 for idx, value in enumerate(time_values)}
    merged["TimeSlotKey"] = merged["Half Hour"].map(time_lookup)

    numeric_map = {
        "Calls Offered": "CallsOffered",
        "Calls Handled": "CallsAnswered",
        "NG Calls Abandoned": "CallsAbandonedAfterThreshold",
        "Abandoned Within Service Level": "CallsAbandonedWithinThreshold",
        "RONA": "RONA",
        "Overflow Out": "OverflowOut",
        "Calls Routed Non Agent": "CallsRoutedNonAgent",
        "Short Calls": "ShortCalls",
        "Service Level Calls": "ServiceLevelCalls",
        "Service Level Calls Offered": "ServiceLevelCallsOffered",
        "WaitTime": "WaitTimeSeconds",
        "Answer Wait Time": "AnswerWaitTimeSeconds",
        "HandleTime": "HandleTimeSeconds",
        "Max Wait Time": "MaxWaitTimeSeconds",
        "Avg Speed To Answer": "SourceAvgSpeedToAnswerSeconds",
        "Avg Handle Time": "SourceAvgHandleTimeSeconds",
        "Service Level": "SourceServiceLevelPct",
    }
    for source_column, output_column in numeric_map.items():
        merged[output_column] = pd.to_numeric(merged[source_column], errors="coerce").fillna(0)

    int_columns = [
        "CallsOffered",
        "CallsAnswered",
        "CallsAbandonedAfterThreshold",
        "CallsAbandonedWithinThreshold",
        "RONA",
        "OverflowOut",
        "CallsRoutedNonAgent",
        "ShortCalls",
        "ServiceLevelCalls",
        "ServiceLevelCallsOffered",
        "WaitTimeSeconds",
        "AnswerWaitTimeSeconds",
        "HandleTimeSeconds",
        "MaxWaitTimeSeconds",
        "SourceAvgSpeedToAnswerSeconds",
        "SourceAvgHandleTimeSeconds",
    ]
    for column in int_columns:
        merged[column] = merged[column].round(0).astype("int64")

    fact = merged[
        [
            "DateKey",
            "TimeSlotKey",
            "CallTypeKey",
            "TeamKey",
            "CallGroupKey",
            "QueueKey",
            "FunctionKey",
            "EmergencyClassKey",
            "WorkingHoursKey",
            *int_columns,
            "SourceServiceLevelPct",
        ]
    ].copy()
    fact.insert(0, "TelephonyRecordKey", range(1, len(fact) + 1))

    date_range = pd.date_range(merged["Call Date"].min(), merged["Call Date"].max(), freq="D")
    dim_date = pd.DataFrame({"Date": date_range})
    dim_date["DateKey"] = dim_date["Date"].dt.strftime("%Y%m%d").astype(int)
    dim_date["MonthStart"] = dim_date["Date"].dt.to_period("M").dt.to_timestamp()
    dim_date["Year"] = dim_date["Date"].dt.year
    dim_date["QuarterNumber"] = dim_date["Date"].dt.quarter
    dim_date["YearQuarter"] = "Q" + dim_date["QuarterNumber"].astype(str) + " " + dim_date["Year"].astype(str)
    dim_date["MonthNumber"] = dim_date["Date"].dt.month
    dim_date["MonthName"] = dim_date["Date"].dt.month_name()
    dim_date["MonthShort"] = dim_date["Date"].dt.strftime("%b")
    dim_date["YearMonth"] = dim_date["Date"].dt.strftime("%Y-%m")
    dim_date["YearMonthSort"] = dim_date["Date"].dt.strftime("%Y%m").astype(int)
    dim_date["WeekStart"] = dim_date["Date"] - pd.to_timedelta(dim_date["Date"].dt.weekday, unit="D")
    dim_date["ISOWeekNumber"] = dim_date["Date"].dt.isocalendar().week.astype(int)
    dim_date["DayOfWeekNumber"] = dim_date["Date"].dt.weekday + 1
    dim_date["DayOfWeekName"] = dim_date["Date"].dt.day_name()
    dim_date["IsWeekend"] = dim_date["DayOfWeekNumber"].isin([6, 7])
    dim_date["IsWorkingDay"] = ~dim_date["IsWeekend"]
    dim_date["DateLabel"] = dim_date["Date"].dt.strftime("%d %b %Y")
    dim_date = dim_date[
        [
            "DateKey",
            "Date",
            "DateLabel",
            "Year",
            "QuarterNumber",
            "YearQuarter",
            "MonthStart",
            "MonthNumber",
            "MonthName",
            "MonthShort",
            "YearMonth",
            "YearMonthSort",
            "WeekStart",
            "ISOWeekNumber",
            "DayOfWeekNumber",
            "DayOfWeekName",
            "IsWeekend",
            "IsWorkingDay",
        ]
    ]

    dim_time_rows = []
    for time_value, key in time_lookup.items():
        hour, minute, _ = [int(part) for part in time_value.split(":")]
        hour_label = f"{hour:02d}:00"
        if 0 <= hour < 6:
            band = "Night"
        elif 6 <= hour < 9:
            band = "Morning Peak"
        elif 9 <= hour < 17:
            band = "Core Day"
        elif 17 <= hour < 20:
            band = "Evening Peak"
        else:
            band = "Late Evening"
        dim_time_rows.append(
            {
                "TimeSlotKey": key,
                "HalfHour": time_value,
                "Hour": hour,
                "HourLabel": hour_label,
                "Minute": minute,
                "TimeBand": band,
                "TimeSlotSort": hour * 60 + minute,
            }
        )
    dim_time = pd.DataFrame(dim_time_rows)

    dim_call_type = call_type_map.copy()
    dim_call_type["Queue"] = dim_call_type["Call Type"]
    dim_call_type["Emergency Classification"] = dim_call_type.apply(
        lambda row: "Emergency"
        if row["Team"] == "Emergency"
        or "Emergency" in row["Call Type"]
        or row["Call Type"] == "False Alarm"
        else "Non-Emergency",
        axis=1,
    )
    dim_call_type["CallTypeKey"] = dim_call_type["Call Type"].map(call_type_lookup)
    dim_call_type["TeamKey"] = dim_call_type["Team"].map(team_lookup)
    dim_call_type["CallGroupKey"] = dim_call_type["Call Group"].map(call_group_lookup)
    dim_call_type["QueueKey"] = dim_call_type["Queue"].map(queue_lookup)
    dim_call_type["FunctionKey"] = dim_call_type["Function"].map(function_lookup)
    dim_call_type["EmergencyClassKey"] = dim_call_type["Emergency Classification"].map(emergency_lookup)
    dim_call_type["IsEmergency"] = dim_call_type["Emergency Classification"].eq("Emergency")
    dim_call_type = dim_call_type[
        [
            "CallTypeKey",
            "Call Type",
            "Queue",
            "TeamKey",
            "Team",
            "CallGroupKey",
            "Call Group",
            "FunctionKey",
            "Function",
            "EmergencyClassKey",
            "Emergency Classification",
            "IsEmergency",
        ]
    ]

    dim_team = pd.DataFrame(
        [{"TeamKey": key, "Team": value} for value, key in sorted(team_lookup.items(), key=lambda item: item[1])]
    )
    dim_call_group = pd.DataFrame(
        [
            {"CallGroupKey": key, "Call Group": value}
            for value, key in sorted(call_group_lookup.items(), key=lambda item: item[1])
        ]
    )
    dim_queue = pd.DataFrame(
        [{"QueueKey": key, "Queue": value} for value, key in sorted(queue_lookup.items(), key=lambda item: item[1])]
    )
    dim_function = pd.DataFrame(
        [
            {"FunctionKey": key, "Function": value}
            for value, key in sorted(function_lookup.items(), key=lambda item: item[1])
        ]
    )
    dim_emergency = pd.DataFrame(
        [
            {"EmergencyClassKey": 1, "Emergency Classification": "Emergency", "IsEmergency": True},
            {"EmergencyClassKey": 2, "Emergency Classification": "Non-Emergency", "IsEmergency": False},
        ]
    )
    dim_working_hours = pd.DataFrame(
        [
            {"WorkingHoursKey": 0, "Working Hours": "Out of Hours", "IsInHours": False},
            {"WorkingHoursKey": 1, "Working Hours": "In Hours", "IsInHours": True},
        ]
    )
    reference_agent_team = agent_team_map.rename(
        columns={"eGain Agent Team": "AgentTeamSource", "Team": "MappedTeam"}
    )

    notes = pd.DataFrame(
        [
            {
                "NoteKey": 1,
                "Area": "Source Grain",
                "Note": "The main extract is aggregated by date, half-hour and call type; it is not individual call-level data.",
            },
            {
                "NoteKey": 2,
                "Area": "Excluded Rows",
                "Note": f"{len(invalid_date_rows)} source rows have no call date and Source = BCM Calls; they are excluded from the star-schema fact so the Calendar relationship remains valid.",
            },
            {
                "NoteKey": 3,
                "Area": "Queue",
                "Note": "No separate queue column exists; Queue is modelled from Call Type, which is the closest available service-line routing field.",
            },
            {
                "NoteKey": 4,
                "Area": "Agent",
                "Note": "The fact extract does not include agent, extension, logged-in time or agent team keys. The supplied agent team mapping is retained as an unconnected reference table.",
            },
            {
                "NoteKey": 5,
                "Area": "Unavailable KPIs",
                "Note": "Transferred calls, hold time, after-call work, repeat calls and logged-in time are not present in the source extract; placeholder measures return blank.",
            },
            {
                "NoteKey": 6,
                "Area": "Durations",
                "Note": "Duration fields are treated as seconds. HandleTime is used for handling time and as the closest available talk-time proxy.",
            },
            {
                "NoteKey": 7,
                "Area": "Target",
                "Note": "The service-level target measure is set to 90 percent as a configurable reporting assumption because no target table is provided.",
            },
        ]
    )

    write_csv(fact, "fact_telephony.csv")
    write_csv(dim_date, "dim_date.csv")
    write_csv(dim_time, "dim_time_slot.csv")
    write_csv(dim_call_type, "dim_call_type.csv")
    write_csv(dim_team, "dim_team.csv")
    write_csv(dim_call_group, "dim_call_group.csv")
    write_csv(dim_queue, "dim_queue.csv")
    write_csv(dim_function, "dim_function.csv")
    write_csv(dim_emergency, "dim_emergency_classification.csv")
    write_csv(dim_working_hours, "dim_working_hours.csv")
    write_csv(reference_agent_team, "reference_agent_team_mapping.csv")
    write_csv(notes, "model_notes.csv")

    if not invalid_date_rows.empty:
        invalid_date_rows.to_csv(CURATED_DIR / "excluded_no_date_rows.csv", index=False, encoding="utf-8")

    return {
        "source_rows": source_row_count,
        "fact_rows": len(fact),
        "excluded_rows": len(invalid_date_rows),
        "date_min": str(dim_date["Date"].min().date()),
        "date_max": str(dim_date["Date"].max().date()),
        "call_type_count": len(dim_call_type),
        "team_count": len(dim_team),
    }


def tmdl_table(
    table_name: str,
    columns: list[dict[str, object]],
    csv_file: str,
    measures: list[dict[str, str]] | None = None,
    table_props: list[str] | None = None,
) -> str:
    lines: list[str] = [f"table '{table_name}'", f"\tlineageTag: {stable_guid('table-' + table_name)}", ""]
    if table_props:
        for prop in table_props:
            lines.insert(1, f"\t{prop}")
    for measure in measures or []:
        expression = measure["expression"]
        if "\n" in expression:
            lines.append(f"\tmeasure '{measure['name']}' =")
            for expr_line in expression.splitlines():
                lines.append(f"\t\t\t{expr_line}")
        else:
            lines.append(f"\tmeasure '{measure['name']}' = {expression}")
        if measure.get("format"):
            lines.append(f"\t\tformatString: {measure['format']}")
        if measure.get("folder"):
            lines.append(f"\t\tdisplayFolder: {measure['folder']}")
        lines.append(f"\t\tlineageTag: {stable_guid('measure-' + table_name + '-' + measure['name'])}")
        lines.append("")

    for column in columns:
        lines.append(f"\tcolumn '{column['name']}'")
        lines.append(f"\t\tdataType: {column['type']}")
        if column.get("format"):
            lines.append(f"\t\tformatString: {column['format']}")
        if column.get("hidden"):
            lines.append("\t\tisHidden")
        if column.get("key"):
            lines.append("\t\tisKey")
        if column.get("summarize"):
            lines.append(f"\t\tsummarizeBy: {column['summarize']}")
        else:
            lines.append("\t\tsummarizeBy: none")
        if column.get("source"):
            lines.append(f"\t\tsourceColumn: {column['source']}")
        if column.get("sort"):
            lines.append(f"\t\tsortByColumn: '{column['sort']}'")
        lines.append(f"\t\tlineageTag: {stable_guid('column-' + table_name + '-' + column['name'])}")
        if column.get("date_annotation"):
            lines.append("")
            lines.append("\t\tannotation UnderlyingDateTimeDataType = Date")
        lines.append("")

    type_map = {
        "int64": "Int64.Type",
        "double": "type number",
        "string": "type text",
        "boolean": "type logical",
        "dateTime": "type date",
    }
    transform_lines = []
    for column in columns:
        if not column.get("source"):
            continue
        transform_lines.append(f'{{"{column["source"]}", {type_map[column["type"]]}}}')

    lines.append(f"\tpartition '{table_name}' = m")
    lines.append("\t\tmode: import")
    lines.append("\t\tsource =")
    lines.append("\t\t\t\tlet")
    lines.append(
        f'\t\t\t\t\tSource = Csv.Document(File.Contents(CuratedDataFolder & "\\{csv_file}"), [Delimiter=",", Columns={len(transform_lines)}, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),'
    )
    lines.append("\t\t\t\t\tHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),")
    lines.append("\t\t\t\t\tTyped = Table.TransformColumnTypes(")
    lines.append("\t\t\t\t\t\tHeaders,")
    lines.append("\t\t\t\t\t\t{")
    for idx, item in enumerate(transform_lines):
        suffix = "," if idx < len(transform_lines) - 1 else ""
        lines.append(f"\t\t\t\t\t\t\t{item}{suffix}")
    lines.append("\t\t\t\t\t\t},")
    lines.append('\t\t\t\t\t\t"en-GB"')
    lines.append("\t\t\t\t\t)")
    lines.append("\t\t\t\tin")
    lines.append("\t\t\t\t\tTyped")
    lines.append("")
    return "\n".join(lines) + "\n"


def measure(name: str, expression: str, fmt: str = "", folder: str = "") -> dict[str, str]:
    return {"name": name, "expression": expression.strip("\n"), "format": fmt, "folder": folder}


def build_semantic_model() -> None:
    TABLES_DIR.mkdir(parents=True, exist_ok=True)

    expressions = f'''expression ProjectRoot = "{ROOT}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
\tlineageTag: {stable_guid("expression-ProjectRoot")}

\tannotation PBI_ResultType = Text

expression CuratedDataFolder = ProjectRoot & "\\Curated Data"
\tlineageTag: {stable_guid("expression-CuratedDataFolder")}

'''
    (DEFINITION_DIR / "expressions.tmdl").write_text(expressions, encoding="utf-8")

    fact_columns = [
        {"name": "Telephony Record Key", "source": "TelephonyRecordKey", "type": "int64", "hidden": True, "key": True},
        {"name": "Date Key", "source": "DateKey", "type": "int64", "hidden": True},
        {"name": "Time Slot Key", "source": "TimeSlotKey", "type": "int64", "hidden": True},
        {"name": "Call Type Key", "source": "CallTypeKey", "type": "int64", "hidden": True},
        {"name": "Team Key", "source": "TeamKey", "type": "int64", "hidden": True},
        {"name": "Call Group Key", "source": "CallGroupKey", "type": "int64", "hidden": True},
        {"name": "Queue Key", "source": "QueueKey", "type": "int64", "hidden": True},
        {"name": "Function Key", "source": "FunctionKey", "type": "int64", "hidden": True},
        {"name": "Emergency Class Key", "source": "EmergencyClassKey", "type": "int64", "hidden": True},
        {"name": "Working Hours Key", "source": "WorkingHoursKey", "type": "int64", "hidden": True},
        {"name": "Calls Offered Raw", "source": "CallsOffered", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Calls Answered Raw", "source": "CallsAnswered", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Calls Abandoned After Threshold Raw", "source": "CallsAbandonedAfterThreshold", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Calls Abandoned Within Threshold Raw", "source": "CallsAbandonedWithinThreshold", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "RONA Raw", "source": "RONA", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Overflow Out Raw", "source": "OverflowOut", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Calls Routed Non Agent Raw", "source": "CallsRoutedNonAgent", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Short Calls Raw", "source": "ShortCalls", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Service Level Calls Raw", "source": "ServiceLevelCalls", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Service Level Calls Offered Raw", "source": "ServiceLevelCallsOffered", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Wait Time Seconds Raw", "source": "WaitTimeSeconds", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Answer Wait Time Seconds Raw", "source": "AnswerWaitTimeSeconds", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Handle Time Seconds Raw", "source": "HandleTimeSeconds", "type": "int64", "hidden": True, "summarize": "sum"},
        {"name": "Max Wait Time Seconds Raw", "source": "MaxWaitTimeSeconds", "type": "int64", "hidden": True, "summarize": "max"},
        {"name": "Source Avg Speed To Answer Seconds Raw", "source": "SourceAvgSpeedToAnswerSeconds", "type": "int64", "hidden": True, "summarize": "none"},
        {"name": "Source Avg Handle Time Seconds Raw", "source": "SourceAvgHandleTimeSeconds", "type": "int64", "hidden": True, "summarize": "none"},
        {"name": "Source Service Level % Raw", "source": "SourceServiceLevelPct", "type": "double", "hidden": True, "summarize": "none"},
    ]

    fact_measures = [
        measure("Calls Offered", "SUM('Fact Telephony'[Calls Offered Raw])", "#,##0", "01 Volume"),
        measure("Service Level Calls Offered", "SUM('Fact Telephony'[Service Level Calls Offered Raw])", "#,##0", "01 Volume"),
        measure("Calls Answered", "SUM('Fact Telephony'[Calls Answered Raw])", "#,##0", "01 Volume"),
        measure("Calls Abandoned Within Service Threshold", "SUM('Fact Telephony'[Calls Abandoned Within Threshold Raw])", "#,##0", "01 Volume"),
        measure("Calls Abandoned After Service Threshold", "SUM('Fact Telephony'[Calls Abandoned After Threshold Raw])", "#,##0", "01 Volume"),
        measure("Calls Abandoned", "[Calls Abandoned Within Service Threshold] + [Calls Abandoned After Service Threshold]", "#,##0", "01 Volume"),
        measure("Calls Not Answered", "[Calls Offered] - [Calls Answered]", "#,##0", "01 Volume"),
        measure("Calls Missed", "SUM('Fact Telephony'[RONA Raw])", "#,##0", "01 Volume"),
        measure("Calls Transferred", "BLANK()", "#,##0", "09 Unavailable Source Fields"),
        measure("Calls Overflowed", "SUM('Fact Telephony'[Overflow Out Raw])", "#,##0", "01 Volume"),
        measure("Calls Routed Non Agent", "SUM('Fact Telephony'[Calls Routed Non Agent Raw])", "#,##0", "01 Volume"),
        measure("Short Calls", "SUM('Fact Telephony'[Short Calls Raw])", "#,##0", "01 Volume"),
        measure("Service Level Calls", "SUM('Fact Telephony'[Service Level Calls Raw])", "#,##0", "01 Volume"),
        measure("Answer Rate %", "DIVIDE([Calls Answered], [Calls Offered])", "0.0%", "02 Rates"),
        measure("Abandonment Rate %", "DIVIDE([Calls Abandoned], [Calls Offered])", "0.0%", "02 Rates"),
        measure("Standards of Service %", "DIVIDE([Service Level Calls], [Service Level Calls Offered])", "0.0%", "02 Rates"),
        measure(
            "Emergency Standards of Service %",
            'CALCULATE(\n    [Standards of Service %],\n    KEEPFILTERS(\'Dim Emergency Classification\'[Emergency Classification] = "Emergency")\n)',
            "0.0%",
            "02 Rates",
        ),
        measure(
            "Emergency Answer Rate %",
            'CALCULATE(\n    [Answer Rate %],\n    KEEPFILTERS(\'Dim Emergency Classification\'[Emergency Classification] = "Emergency")\n)',
            "0.0%",
            "02 Rates",
        ),
        measure(
            "Emergency Abandonment Rate %",
            'CALCULATE(\n    [Abandonment Rate %],\n    KEEPFILTERS(\'Dim Emergency Classification\'[Emergency Classification] = "Emergency")\n)',
            "0.0%",
            "02 Rates",
        ),
        measure("Service Level Target %", "0.90", "0.0%", "02 Rates"),
        measure("Standards of Service Variance to Target", "[Standards of Service %] - [Service Level Target %]", "+0.0%;-0.0%;0.0%", "02 Rates"),
        measure(
            "Below Service Target Flag",
            "IF(NOT ISBLANK([Standards of Service %]) && [Standards of Service %] < [Service Level Target %], 1, 0)",
            "0",
            "02 Rates",
        ),
        measure("Total Wait Time Seconds", "SUM('Fact Telephony'[Wait Time Seconds Raw])", "#,##0", "03 Wait and Handle Time"),
        measure("Total Answer Wait Time Seconds", "SUM('Fact Telephony'[Answer Wait Time Seconds Raw])", "#,##0", "03 Wait and Handle Time"),
        measure("Average Speed To Answer", "DIVIDE([Total Answer Wait Time Seconds], [Calls Answered])", "#,##0.0", "03 Wait and Handle Time"),
        measure("Average Wait Time", "DIVIDE([Total Wait Time Seconds], [Calls Offered])", "#,##0.0", "03 Wait and Handle Time"),
        measure("Longest Wait Time", "MAX('Fact Telephony'[Max Wait Time Seconds Raw])", "#,##0", "03 Wait and Handle Time"),
        measure("Total Handling Time Seconds", "SUM('Fact Telephony'[Handle Time Seconds Raw])", "#,##0", "03 Wait and Handle Time"),
        measure("Average Handling Time", "DIVIDE([Total Handling Time Seconds], [Calls Answered])", "#,##0.0", "03 Wait and Handle Time"),
        measure("Average Talk Time", "[Average Handling Time]", "#,##0.0", "03 Wait and Handle Time"),
        measure("Total Talk Time", "[Total Handling Time Seconds]", "#,##0", "03 Wait and Handle Time"),
        measure("Average Hold Time", "BLANK()", "#,##0.0", "09 Unavailable Source Fields"),
        measure("After-Call Work Time", "BLANK()", "#,##0", "09 Unavailable Source Fields"),
        measure("Calls Below Service Threshold", "[Service Level Calls]", "#,##0", "04 Threshold Bands"),
        measure("Calls Above Service Threshold", "MAX(0, [Service Level Calls Offered] - [Service Level Calls])", "#,##0", "04 Threshold Bands"),
        measure("Answered Within Target", "[Service Level Calls]", "#,##0", "04 Threshold Bands"),
        measure("Answered After Target", "MAX(0, [Calls Answered] - [Service Level Calls])", "#,##0", "04 Threshold Bands"),
        measure("Abandoned Within Target", "[Calls Abandoned Within Service Threshold]", "#,##0", "04 Threshold Bands"),
        measure("Abandoned After Target", "[Calls Abandoned After Service Threshold]", "#,##0", "04 Threshold Bands"),
        measure("Active Days", "DISTINCTCOUNT('Fact Telephony'[Date Key])", "#,##0", "05 Productivity"),
        measure(
            "Active Working Days",
            "CALCULATE(\n    DISTINCTCOUNT('Fact Telephony'[Date Key]),\n    KEEPFILTERS('Dim Date'[Is Working Day] = TRUE())\n)",
            "#,##0",
            "05 Productivity",
        ),
        measure("Calls Offered Per Working Day", "DIVIDE([Calls Offered], [Active Working Days])", "#,##0.0", "05 Productivity"),
        measure("Calls Answered Per Working Day", "DIVIDE([Calls Answered], [Active Working Days])", "#,##0.0", "05 Productivity"),
        measure("Average Calls Per Working Day", "[Calls Offered Per Working Day]", "#,##0.0", "05 Productivity"),
        measure("Calls Answered Per Agent", "BLANK()", "#,##0.0", "09 Unavailable Source Fields"),
        measure("Total Call Time", "[Total Handling Time Seconds]", "#,##0", "05 Productivity"),
        measure("Transfer Rate %", "BLANK()", "0.0%", "09 Unavailable Source Fields"),
        measure("Latest Data Date", "CALCULATE(MAX('Dim Date'[Date]), REMOVEFILTERS('Dim Date'))", "Long Date", "06 Latest Snapshot"),
        measure("Latest Month Label", "FORMAT(CALCULATE(MAX('Dim Date'[Month Start]), REMOVEFILTERS('Dim Date')), \"mmmm yyyy\")", "", "06 Latest Snapshot"),
        measure(
            "Latest Month Calls Offered",
            "VAR LatestMonth = CALCULATE(MAX('Dim Date'[Month Start]), REMOVEFILTERS('Dim Date'))\nRETURN\n    CALCULATE([Calls Offered], REMOVEFILTERS('Dim Date'), 'Dim Date'[Month Start] = LatestMonth)",
            "#,##0",
            "06 Latest Snapshot",
        ),
        measure(
            "Previous Month Calls Offered",
            "VAR LatestMonth = CALCULATE(MAX('Dim Date'[Month Start]), REMOVEFILTERS('Dim Date'))\nRETURN\n    CALCULATE([Calls Offered], REMOVEFILTERS('Dim Date'), 'Dim Date'[Month Start] = EDATE(LatestMonth, -1))",
            "#,##0",
            "06 Latest Snapshot",
        ),
        measure("Calls Offered MoM Change %", "DIVIDE([Latest Month Calls Offered] - [Previous Month Calls Offered], [Previous Month Calls Offered])", "+0.0%;-0.0%;0.0%", "06 Latest Snapshot"),
        measure(
            "Latest Month Standards of Service %",
            "VAR LatestMonth = CALCULATE(MAX('Dim Date'[Month Start]), REMOVEFILTERS('Dim Date'))\nRETURN\n    CALCULATE([Standards of Service %], REMOVEFILTERS('Dim Date'), 'Dim Date'[Month Start] = LatestMonth)",
            "0.0%",
            "06 Latest Snapshot",
        ),
        measure(
            "Previous Month Standards of Service %",
            "VAR LatestMonth = CALCULATE(MAX('Dim Date'[Month Start]), REMOVEFILTERS('Dim Date'))\nRETURN\n    CALCULATE([Standards of Service %], REMOVEFILTERS('Dim Date'), 'Dim Date'[Month Start] = EDATE(LatestMonth, -1))",
            "0.0%",
            "06 Latest Snapshot",
        ),
        measure("Standards of Service MoM Change", "[Latest Month Standards of Service %] - [Previous Month Standards of Service %]", "+0.0%;-0.0%;0.0%", "06 Latest Snapshot"),
        measure(
            "Service Status",
            "SWITCH(\n    TRUE(),\n    ISBLANK([Standards of Service %]), BLANK(),\n    [Standards of Service %] >= [Service Level Target %], \"On target\",\n    \"Below target\"\n)",
            "",
            "07 Narrative",
        ),
        measure(
            "Highest Abandonment Queue",
            "VAR Candidates =\n    ADDCOLUMNS(\n        ALL('Dim Queue'[Queue]),\n        \"__Abandoned\", [Calls Abandoned]\n    )\nVAR Ranked =\n    TOPN(1, Candidates, [__Abandoned], DESC, 'Dim Queue'[Queue], ASC)\nRETURN\n    MAXX(Ranked, 'Dim Queue'[Queue])",
            "",
            "07 Narrative",
        ),
        measure(
            "Highest Pressure Hour",
            "VAR Candidates =\n    ADDCOLUMNS(\n        ALL('Dim Time Slot'[Hour Label]),\n        \"__Offered\", [Calls Offered]\n    )\nVAR Ranked =\n    TOPN(1, Candidates, [__Offered], DESC, 'Dim Time Slot'[Hour Label], ASC)\nRETURN\n    MAXX(Ranked, 'Dim Time Slot'[Hour Label])",
            "",
            "07 Narrative",
        ),
        measure(
            "Headline Narrative",
            "VAR StatusText = [Service Status]\nVAR LatestMonth = [Latest Month Label]\nVAR Sos = [Latest Month Standards of Service %]\nVAR QueueName = [Highest Abandonment Queue]\nVAR PeakHour = [Highest Pressure Hour]\nRETURN\n    IF(\n        ISBLANK(Sos),\n        BLANK(),\n        \"As of \" & LatestMonth & \", Standards of Service is \" & FORMAT(Sos, \"0.0%\") & \" and currently \" & LOWER(StatusText) & \". The highest abandonment queue is \" & COALESCE(QueueName, \"not available\") & \", with demand peaking around \" & COALESCE(PeakHour, \"not available\") & \".\"\n    )",
            "",
            "07 Narrative",
        ),
    ]

    tables: dict[str, str] = {}
    tables["Fact Telephony"] = tmdl_table("Fact Telephony", fact_columns, "fact_telephony.csv", fact_measures)

    tables["Dim Date"] = tmdl_table(
        "Dim Date",
        [
            {"name": "Date Key", "source": "DateKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Date", "source": "Date", "type": "dateTime", "format": "Long Date", "date_annotation": True},
            {"name": "Date Label", "source": "DateLabel", "type": "string"},
            {"name": "Year", "source": "Year", "type": "int64", "summarize": "none"},
            {"name": "Quarter Number", "source": "QuarterNumber", "type": "int64", "hidden": True},
            {"name": "Year Quarter", "source": "YearQuarter", "type": "string"},
            {"name": "Month Start", "source": "MonthStart", "type": "dateTime", "format": "Long Date", "hidden": True, "date_annotation": True},
            {"name": "Month Number", "source": "MonthNumber", "type": "int64", "hidden": True},
            {"name": "Month Name", "source": "MonthName", "type": "string", "sort": "Month Number"},
            {"name": "Month Short", "source": "MonthShort", "type": "string", "hidden": True, "sort": "Month Number"},
            {"name": "Year Month", "source": "YearMonth", "type": "string", "sort": "Year Month Sort"},
            {"name": "Year Month Sort", "source": "YearMonthSort", "type": "int64", "hidden": True},
            {"name": "Week Start", "source": "WeekStart", "type": "dateTime", "format": "Long Date", "date_annotation": True},
            {"name": "ISO Week Number", "source": "ISOWeekNumber", "type": "int64", "summarize": "none"},
            {"name": "Day Of Week Number", "source": "DayOfWeekNumber", "type": "int64", "hidden": True},
            {"name": "Day Of Week", "source": "DayOfWeekName", "type": "string", "sort": "Day Of Week Number"},
            {"name": "Is Weekend", "source": "IsWeekend", "type": "boolean", "hidden": True},
            {"name": "Is Working Day", "source": "IsWorkingDay", "type": "boolean"},
        ],
        "dim_date.csv",
    )
    tables["Dim Time Slot"] = tmdl_table(
        "Dim Time Slot",
        [
            {"name": "Time Slot Key", "source": "TimeSlotKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Half Hour", "source": "HalfHour", "type": "string", "sort": "Time Slot Sort"},
            {"name": "Hour", "source": "Hour", "type": "int64", "hidden": True},
            {"name": "Hour Label", "source": "HourLabel", "type": "string", "sort": "Hour"},
            {"name": "Minute", "source": "Minute", "type": "int64", "hidden": True},
            {"name": "Time Band", "source": "TimeBand", "type": "string"},
            {"name": "Time Slot Sort", "source": "TimeSlotSort", "type": "int64", "hidden": True},
        ],
        "dim_time_slot.csv",
    )
    tables["Dim Call Type"] = tmdl_table(
        "Dim Call Type",
        [
            {"name": "Call Type Key", "source": "CallTypeKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Call Type", "source": "Call Type", "type": "string"},
            {"name": "Queue", "source": "Queue", "type": "string"},
            {"name": "Team Key", "source": "TeamKey", "type": "int64", "hidden": True},
            {"name": "Team", "source": "Team", "type": "string"},
            {"name": "Call Group Key", "source": "CallGroupKey", "type": "int64", "hidden": True},
            {"name": "Call Group", "source": "Call Group", "type": "string"},
            {"name": "Function Key", "source": "FunctionKey", "type": "int64", "hidden": True},
            {"name": "Function", "source": "Function", "type": "string"},
            {"name": "Emergency Class Key", "source": "EmergencyClassKey", "type": "int64", "hidden": True},
            {"name": "Emergency Classification", "source": "Emergency Classification", "type": "string"},
            {"name": "Is Emergency", "source": "IsEmergency", "type": "boolean"},
        ],
        "dim_call_type.csv",
    )
    tables["Dim Team"] = tmdl_table(
        "Dim Team",
        [
            {"name": "Team Key", "source": "TeamKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Team", "source": "Team", "type": "string"},
        ],
        "dim_team.csv",
    )
    tables["Dim Call Group"] = tmdl_table(
        "Dim Call Group",
        [
            {"name": "Call Group Key", "source": "CallGroupKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Call Group", "source": "Call Group", "type": "string"},
        ],
        "dim_call_group.csv",
    )
    tables["Dim Queue"] = tmdl_table(
        "Dim Queue",
        [
            {"name": "Queue Key", "source": "QueueKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Queue", "source": "Queue", "type": "string"},
        ],
        "dim_queue.csv",
    )
    tables["Dim Function"] = tmdl_table(
        "Dim Function",
        [
            {"name": "Function Key", "source": "FunctionKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Function", "source": "Function", "type": "string"},
        ],
        "dim_function.csv",
    )
    tables["Dim Emergency Classification"] = tmdl_table(
        "Dim Emergency Classification",
        [
            {"name": "Emergency Class Key", "source": "EmergencyClassKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Emergency Classification", "source": "Emergency Classification", "type": "string"},
            {"name": "Is Emergency", "source": "IsEmergency", "type": "boolean"},
        ],
        "dim_emergency_classification.csv",
    )
    tables["Dim Working Hours"] = tmdl_table(
        "Dim Working Hours",
        [
            {"name": "Working Hours Key", "source": "WorkingHoursKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Working Hours", "source": "Working Hours", "type": "string"},
            {"name": "Is In Hours", "source": "IsInHours", "type": "boolean"},
        ],
        "dim_working_hours.csv",
    )
    tables["Model Notes"] = tmdl_table(
        "Model Notes",
        [
            {"name": "Note Key", "source": "NoteKey", "type": "int64", "hidden": True, "key": True},
            {"name": "Area", "source": "Area", "type": "string"},
            {"name": "Note", "source": "Note", "type": "string"},
        ],
        "model_notes.csv",
    )
    tables["Reference Agent Team Mapping"] = tmdl_table(
        "Reference Agent Team Mapping",
        [
            {"name": "Agent Team Source", "source": "AgentTeamSource", "type": "string"},
            {"name": "Mapped Team", "source": "MappedTeam", "type": "string"},
        ],
        "reference_agent_team_mapping.csv",
        table_props=["isHidden"],
    )

    for table_name, content in tables.items():
        (TABLES_DIR / f"{table_name}.tmdl").write_text(content, encoding="utf-8")

    rels = [
        ("Fact Telephony", "Date Key", "Dim Date", "Date Key"),
        ("Fact Telephony", "Time Slot Key", "Dim Time Slot", "Time Slot Key"),
        ("Fact Telephony", "Call Type Key", "Dim Call Type", "Call Type Key"),
        ("Fact Telephony", "Team Key", "Dim Team", "Team Key"),
        ("Fact Telephony", "Call Group Key", "Dim Call Group", "Call Group Key"),
        ("Fact Telephony", "Queue Key", "Dim Queue", "Queue Key"),
        ("Fact Telephony", "Function Key", "Dim Function", "Function Key"),
        ("Fact Telephony", "Emergency Class Key", "Dim Emergency Classification", "Emergency Class Key"),
        ("Fact Telephony", "Working Hours Key", "Dim Working Hours", "Working Hours Key"),
    ]
    rel_lines = []
    for idx, (from_table, from_column, to_table, to_column) in enumerate(rels, start=1):
        rel_lines.append(f"relationship {stable_guid('relationship-' + str(idx) + from_table + to_table)}")
        rel_lines.append(f"\tfromColumn: '{from_table}'.'{from_column}'")
        rel_lines.append(f"\ttoColumn: '{to_table}'.'{to_column}'")
        rel_lines.append("")
    (DEFINITION_DIR / "relationships.tmdl").write_text("\n".join(rel_lines), encoding="utf-8")

    model_lines = [
        "model Model",
        "\tculture: en-GB",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tsourceQueryCulture: en-GB",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        'annotation PBI_ProTooling = ["DevMode"]',
        "",
        "ref expression ProjectRoot",
        "ref expression CuratedDataFolder",
        "",
    ]
    for table_name in tables:
        model_lines.append(f"ref table '{table_name}'")
    model_lines.extend(["", "ref cultureInfo en-GB", ""])
    model_path = DEFINITION_DIR / "model.tmdl"
    try:
        model_path.write_text("\n".join(model_lines), encoding="utf-8")
    except PermissionError:
        model_path.with_suffix(".tmdl.generated").write_text("\n".join(model_lines), encoding="utf-8")


def lit(value: object) -> dict[str, dict[str, str]]:
    if isinstance(value, bool):
        text = "true" if value else "false"
    elif isinstance(value, int):
        text = f"{value}L"
    elif isinstance(value, float):
        text = f"{value}D"
    else:
        safe = str(value).replace("'", "''")
        text = f"'{safe}'"
    return {"expr": {"Literal": {"Value": text}}}


def solid(color: str) -> dict[str, object]:
    return {"solid": {"color": lit(color)}}


def source_ref(entity: str) -> dict[str, dict[str, str]]:
    return {"SourceRef": {"Entity": entity}}


def field_column(entity: str, prop: str) -> dict[str, object]:
    return {"Column": {"Expression": source_ref(entity), "Property": prop}}


def field_measure(entity: str, prop: str) -> dict[str, object]:
    return {"Measure": {"Expression": source_ref(entity), "Property": prop}}


def projection(entity: str, prop: str, kind: str = "measure", active: bool = False) -> dict[str, object]:
    field = field_measure(entity, prop) if kind == "measure" else field_column(entity, prop)
    item = {"field": field, "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}
    if active:
        item["active"] = True
    return item


def container_objects(title: str | None = None, background: str = "#FFFFFF", border: bool = False) -> dict[str, object]:
    objects: dict[str, object] = {
        "background": [
            {
                "properties": {
                    "show": lit(True),
                    "color": solid(background),
                    "transparency": lit(4 if background != "#FFFFFF" else 20),
                }
            }
        ],
        "visualHeader": [{"properties": {"show": lit(False), "transparency": lit(100.0)}}],
    }
    if border:
        objects["border"] = [
            {
                "properties": {
                    "show": lit(True),
                    "color": solid("#D7DEE8"),
                    "radius": lit(5.0),
                    "width": lit(1.0),
                }
            }
        ]
    if title:
        objects["title"] = [
            {
                "properties": {
                    "show": lit(True),
                    "text": lit(title),
                    "titleWrap": lit(True),
                    "fontColor": solid("#1F2937"),
                    "alignment": lit("center"),
                    "fontSize": lit(11.0),
                    "bold": lit(True),
                }
            }
        ]
    return objects


def visual_base(name: str, x: float, y: float, width: float, height: float, z: int, visual: dict[str, object]) -> dict[str, object]:
    return {
        "$schema": REPORT_SCHEMA,
        "name": name,
        "position": {
            "x": x,
            "y": y,
            "z": z,
            "height": height,
            "width": width,
            "tabOrder": z,
        },
        "visual": visual,
    }


def textbox(name: str, x: float, y: float, width: float, height: float, z: int, text: str, size: int = 16, color: str = "#FFFFFF", bold: bool = True) -> dict[str, object]:
    return visual_base(
        name,
        x,
        y,
        width,
        height,
        z,
        {
            "visualType": "textbox",
            "objects": {
                "general": [
                    {
                        "properties": {
                            "paragraphs": [
                                {
                                    "textRuns": [
                                        {
                                            "value": text,
                                            "textStyle": {
                                                "fontWeight": "bold" if bold else "normal",
                                                "fontSize": f"{size}pt",
                                                "color": color,
                                            },
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                ]
            },
            "visualContainerObjects": {"background": [{"properties": {"show": lit(False)}}]},
            "drillFilterOtherVisuals": True,
        },
    )


def shape(name: str, x: float, y: float, width: float, height: float, z: int, color: str) -> dict[str, object]:
    return visual_base(
        name,
        x,
        y,
        width,
        height,
        z,
        {
            "visualType": "shape",
            "objects": {
                "shape": [{"properties": {"roundEdge": lit(5)}}],
                "fill": [{"properties": {"show": lit(True), "fillColor": solid(color), "transparency": lit(0.0)}}],
                "outline": [{"properties": {"show": lit(False)}}],
            },
            "drillFilterOtherVisuals": True,
        },
    )


def card(name: str, x: float, y: float, width: float, height: float, z: int, measure_name: str, title: str) -> dict[str, object]:
    visual = {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [projection("Fact Telephony", measure_name)]}}},
        "objects": {
            "categoryLabels": [{"properties": {"show": lit(True), "fontSize": lit(9.0), "color": solid("#64748B")}}],
            "labels": [
                {
                    "properties": {
                        "color": solid("#0F5F73"),
                        "fontSize": lit(19.0),
                        "bold": lit(True),
                        "fontFamily": lit("Segoe UI Semibold"),
                        "labelDisplayUnits": lit(1.0),
                    }
                }
            ],
        },
        "visualContainerObjects": container_objects(title, "#FFFFFF", True),
        "drillFilterOtherVisuals": True,
    }
    return visual_base(name, x, y, width, height, z, visual)


def card_visual(name: str, x: float, y: float, width: float, height: float, z: int, measure_name: str, title: str) -> dict[str, object]:
    visual = {
        "visualType": "cardVisual",
        "query": {"queryState": {"Data": {"projections": [projection("Fact Telephony", measure_name)]}}},
        "objects": {
            "value": [{"properties": {"fontSize": lit(12.0), "textWrap": lit(True), "fontColor": solid("#0F172A")}}],
            "label": [{"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
        },
        "visualContainerObjects": container_objects(title, "#F8FAFC", True),
        "drillFilterOtherVisuals": True,
    }
    return visual_base(name, x, y, width, height, z, visual)


def slicer(name: str, x: float, y: float, width: float, height: float, z: int, entity: str, prop: str, title: str) -> dict[str, object]:
    visual = {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [projection(entity, prop, "column", True)]}}},
        "objects": {
            "data": [{"properties": {"mode": lit("Dropdown")}}],
            "selection": [{"properties": {"singleSelect": lit(False)}}],
        },
        "visualContainerObjects": container_objects(title, "#FFFFFF", True),
        "drillFilterOtherVisuals": True,
    }
    return visual_base(name, x, y, width, height, z, visual)


def chart(
    name: str,
    visual_type: str,
    x: float,
    y: float,
    width: float,
    height: float,
    z: int,
    category_entity: str,
    category_prop: str,
    measures: list[str],
    title: str,
) -> dict[str, object]:
    visual = {
        "visualType": visual_type,
        "query": {
            "queryState": {
                "Category": {"projections": [projection(category_entity, category_prop, "column", True)]},
                "Y": {"projections": [projection("Fact Telephony", item) for item in measures]},
            }
        },
        "objects": {
            "legend": [{"properties": {"show": lit(len(measures) > 1)}}],
            "valueAxis": [{"properties": {"show": lit(True), "fontSize": lit(8.0), "titleFontSize": lit(8.0)}}],
            "categoryAxis": [{"properties": {"fontSize": lit(8.0), "maxMarginFactor": lit(32)}}],
            "dataPoint": [{"properties": {"fill": solid("#0F5F73")}}],
        },
        "visualContainerObjects": container_objects(title, "#FFFFFF", True),
        "drillFilterOtherVisuals": True,
    }
    return visual_base(name, x, y, width, height, z, visual)


def matrix(name: str, x: float, y: float, width: float, height: float, z: int, rows: list[tuple[str, str]], measures: list[str], title: str) -> dict[str, object]:
    visual = {
        "visualType": "pivotTable",
        "query": {
            "queryState": {
                "Rows": {"projections": [projection(entity, prop, "column", idx == 0) for idx, (entity, prop) in enumerate(rows)]},
                "Values": {"projections": [projection("Fact Telephony", item) for item in measures]},
            }
        },
        "objects": {
            "grid": [{"properties": {"outlineColor": solid("#E5E7EB")}}],
            "columnHeaders": [{"properties": {"backColor": solid("#E6F1F4"), "fontColor": solid("#0F172A")}}],
            "values": [{"properties": {"backColorPrimary": solid("#FFFFFF"), "backColorSecondary": solid("#F8FAFC")}}],
        },
        "visualContainerObjects": container_objects(title, "#FFFFFF", True),
        "drillFilterOtherVisuals": True,
    }
    return visual_base(name, x, y, width, height, z, visual)


def create_page(page_name: str, display_name: str, subtitle: str, visuals: list[dict[str, object]]) -> None:
    page_dir = PAGES_DIR / page_name
    visuals_dir = page_dir / "visuals"
    if visuals_dir.exists():
        shutil.rmtree(visuals_dir)
    visuals_dir.mkdir(parents=True, exist_ok=True)
    page_json = {
        "$schema": PAGE_SCHEMA,
        "name": page_name,
        "displayName": display_name,
        "displayOption": "FitToPage",
        "height": 720,
        "width": 1280,
    }
    page_json_path = page_dir / "page.json"
    try:
        page_json_path.write_text(json.dumps(page_json, indent=2), encoding="utf-8")
    except PermissionError:
        page_json_path.with_suffix(".json.generated").write_text(json.dumps(page_json, indent=2), encoding="utf-8")

    header = [
        shape(f"{page_name[:6]}_header_shape", 30, 12, 1220, 58, 1, "#0F5F73"),
        textbox(f"{page_name[:6]}_header_title", 52, 20, 560, 28, 2, display_name, 17, "#FFFFFF", True),
        textbox(f"{page_name[:6]}_header_subtitle", 620, 23, 590, 24, 3, subtitle, 10, "#E6F1F4", False),
    ]
    for visual in header + visuals:
        visual_dir = visuals_dir / visual["name"]
        visual_dir.mkdir(parents=True, exist_ok=True)
        (visual_dir / "visual.json").write_text(json.dumps(visual, indent=2), encoding="utf-8")


def build_report() -> None:
    PAGES_DIR.mkdir(parents=True, exist_ok=True)
    page_ids = [
        "c63853cc17e9decfa742",
        "67b2e9fa01d344c8a101",
        "8d02acbe44f96e2b7730",
        "b37cf93a58de4201c445",
        "e45df22f768b45d08a90",
    ]

    slicers_top = [
        ("sl_date", "Dim Date", "Date", "Date"),
        ("sl_team", "Dim Team", "Team", "Team"),
        ("sl_group", "Dim Call Group", "Call Group", "Call Group"),
        ("sl_queue", "Dim Queue", "Queue", "Queue"),
        ("sl_emerg", "Dim Emergency Classification", "Emergency Classification", "Emergency"),
        ("sl_hours", "Dim Working Hours", "Working Hours", "Hours"),
    ]

    def slicer_row(prefix: str, y: float = 82) -> list[dict[str, object]]:
        width = 190
        gap = 12
        return [
            slicer(f"{prefix}_{sid}", 35 + idx * (width + gap), y, width, 45, 100 + idx, entity, prop, title)
            for idx, (sid, entity, prop, title) in enumerate(slicers_top)
        ]

    page1 = slicer_row("p1") + [
        card("p1_card_offered", 35, 145, 170, 82, 300, "Calls Offered", "Calls offered"),
        card("p1_card_answered", 215, 145, 170, 82, 301, "Calls Answered", "Calls answered"),
        card("p1_card_abandoned", 395, 145, 170, 82, 302, "Calls Abandoned", "Calls abandoned"),
        card("p1_card_answer_rate", 575, 145, 170, 82, 303, "Answer Rate %", "Answer rate"),
        card("p1_card_abandon_rate", 755, 145, 170, 82, 304, "Abandonment Rate %", "Abandonment rate"),
        card("p1_card_sos", 935, 145, 145, 82, 305, "Standards of Service %", "Standards"),
        card("p1_card_emerg", 1090, 145, 160, 82, 306, "Emergency Standards of Service %", "Emergency standards"),
        card_visual("p1_narrative", 35, 238, 1215, 62, 307, "Headline Narrative", "Current service readout"),
        chart("p1_line_volume", "lineChart", 35, 315, 590, 175, 400, "Dim Date", "Year Month", ["Calls Offered", "Calls Answered", "Calls Abandoned"], "Monthly call volumes"),
        chart("p1_line_rates", "lineChart", 645, 315, 605, 175, 401, "Dim Date", "Year Month", ["Standards of Service %", "Answer Rate %", "Abandonment Rate %"], "Monthly standards and rates"),
        chart("p1_bar_group", "clusteredBarChart", 35, 510, 445, 165, 402, "Dim Call Group", "Call Group", ["Calls Offered"], "Demand by call group"),
        matrix("p1_month_matrix", 500, 510, 750, 165, 403, [("Dim Date", "Year Month")], ["Calls Offered", "Calls Answered", "Calls Abandoned", "Standards of Service %"], "Monthly summary"),
    ]
    create_page(page_ids[0], "Overall Standards of Service", "Executive summary, standards trend and key filters", page1)

    page2 = slicer_row("p2") + [
        card("p2_card_sos", 35, 145, 190, 82, 300, "Standards of Service %", "Standards"),
        card("p2_card_target", 235, 145, 190, 82, 301, "Standards of Service Variance to Target", "Vs target"),
        card("p2_card_asa", 435, 145, 190, 82, 302, "Average Speed To Answer", "Avg speed to answer"),
        card("p2_card_aht", 635, 145, 190, 82, 303, "Average Handling Time", "Avg handling time"),
        card("p2_card_wait", 835, 145, 190, 82, 304, "Average Wait Time", "Avg wait time"),
        card("p2_card_long_wait", 1035, 145, 215, 82, 305, "Longest Wait Time", "Longest wait"),
        chart("p2_team_sos", "clusteredBarChart", 35, 245, 380, 190, 400, "Dim Team", "Team", ["Standards of Service %", "Service Level Target %"], "Standards by team"),
        chart("p2_group_abandon", "clusteredBarChart", 435, 245, 380, 190, 401, "Dim Call Group", "Call Group", ["Abandonment Rate %"], "Abandonment rate by call group"),
        chart("p2_month_sos", "lineChart", 835, 245, 415, 190, 402, "Dim Date", "Year Month", ["Standards of Service %", "Service Level Target %"], "Standards trend"),
        chart("p2_day_pressure", "clusteredColumnChart", 35, 455, 380, 180, 403, "Dim Date", "Day Of Week", ["Calls Offered"], "Demand by day of week"),
        chart("p2_hour_pressure", "clusteredColumnChart", 435, 455, 380, 180, 404, "Dim Time Slot", "Hour Label", ["Calls Offered", "Calls Abandoned"], "Demand by hour"),
        matrix("p2_team_matrix", 835, 455, 415, 180, 405, [("Dim Team", "Team"), ("Dim Call Group", "Call Group")], ["Calls Offered", "Calls Answered", "Calls Abandoned", "Standards of Service %", "Average Speed To Answer", "Average Handling Time"], "Team and call group detail"),
    ]
    create_page(page_ids[1], "Team and Call Group Performance", "Compare service levels, demand and handling performance across mapped teams", page2)

    page3 = slicer_row("p3") + [
        card("p3_card_offered", 35, 145, 180, 82, 300, "Calls Offered", "Inbound demand"),
        card("p3_card_abandoned", 225, 145, 180, 82, 301, "Calls Abandoned", "Abandoned calls"),
        card("p3_card_missed", 415, 145, 180, 82, 302, "Calls Missed", "Missed/RONA"),
        card("p3_card_overflow", 605, 145, 180, 82, 303, "Calls Overflowed", "Overflowed"),
        card("p3_card_within", 795, 145, 220, 82, 304, "Abandoned Within Target", "Abandoned within target"),
        card("p3_card_after", 1025, 145, 225, 82, 305, "Abandoned After Target", "Abandoned after target"),
        chart("p3_queue_abandoned", "clusteredBarChart", 35, 245, 400, 190, 400, "Dim Queue", "Queue", ["Calls Abandoned"], "Queues with highest abandonment"),
        chart("p3_abandon_trend", "lineChart", 455, 245, 390, 190, 401, "Dim Date", "Year Month", ["Abandonment Rate %", "Answer Rate %"], "Abandonment and answer trend"),
        chart("p3_hour_deterioration", "clusteredColumnChart", 865, 245, 385, 190, 402, "Dim Time Slot", "Hour Label", ["Calls Offered", "Calls Abandoned"], "Peak periods by hour"),
        chart("p3_thresholds", "clusteredBarChart", 35, 455, 400, 180, 403, "Dim Queue", "Queue", ["Answered Within Target", "Answered After Target"], "Wait-time service bands"),
        chart("p3_emergency", "clusteredColumnChart", 455, 455, 390, 180, 404, "Dim Emergency Classification", "Emergency Classification", ["Standards of Service %", "Answer Rate %", "Abandonment Rate %"], "Emergency vs non-emergency"),
        matrix("p3_queue_matrix", 865, 455, 385, 180, 405, [("Dim Queue", "Queue")], ["Calls Offered", "Calls Answered", "Calls Abandoned", "Calls Missed", "Calls Overflowed", "Average Wait Time", "Longest Wait Time"], "Queue detail"),
    ]
    create_page(page_ids[2], "Queue, Demand and Abandonment Analysis", "Demand pressure, abandonment, wait-time bands and emergency performance", page3)

    page4 = slicer_row("p4") + [
        card_visual("p4_note", 35, 145, 1215, 58, 300, "Headline Narrative", "Operational context"),
        card("p4_card_aht", 35, 220, 190, 82, 301, "Average Handling Time", "Avg handling time"),
        card("p4_card_talk", 235, 220, 190, 82, 302, "Average Talk Time", "Avg talk time proxy"),
        card("p4_card_total_time", 435, 220, 190, 82, 303, "Total Call Time", "Total call time"),
        card("p4_card_working", 635, 220, 190, 82, 304, "Calls Answered Per Working Day", "Answered per working day"),
        card("p4_card_agent", 835, 220, 190, 82, 305, "Calls Answered Per Agent", "Answered per agent"),
        card("p4_card_transfer", 1035, 220, 215, 82, 306, "Transfer Rate %", "Transfer rate"),
        chart("p4_team_aht", "clusteredBarChart", 35, 320, 390, 180, 400, "Dim Team", "Team", ["Average Handling Time", "Average Speed To Answer"], "Handling and answer time by team"),
        chart("p4_calltype_aht", "clusteredBarChart", 445, 320, 390, 180, 401, "Dim Call Type", "Call Type", ["Average Handling Time"], "Long handling call types"),
        chart("p4_hours_workload", "clusteredColumnChart", 855, 320, 395, 180, 402, "Dim Working Hours", "Working Hours", ["Calls Offered", "Calls Answered"], "In-hours and out-of-hours workload"),
        matrix("p4_eff_matrix", 35, 520, 1215, 145, 403, [("Dim Team", "Team"), ("Dim Call Type", "Call Type")], ["Calls Answered", "Average Handling Time", "Average Speed To Answer", "Total Handling Time Seconds", "Calls Offered Per Working Day", "Answer Rate %"], "Operational efficiency detail"),
        textbox("p4_source_limit", 45, 668, 1160, 22, 404, "Agent, extension, hold time, after-call work, repeat-call and logged-in-time fields are not present in the fact extract; placeholder measures are intentionally blank.", 8, "#475569", False),
    ]
    create_page(page_ids[3], "Agent, Handling and Operational Efficiency", "Handling-time and workload analysis from the available aggregate source grain", page4)

    page5 = slicer_row("p5") + [
        card("p5_card_rows", 35, 145, 180, 82, 300, "Active Days", "Active days"),
        card("p5_card_offered", 225, 145, 180, 82, 301, "Calls Offered", "Calls offered"),
        card("p5_card_answered", 415, 145, 180, 82, 302, "Calls Answered", "Calls answered"),
        card("p5_card_abandoned", 605, 145, 180, 82, 303, "Calls Abandoned", "Calls abandoned"),
        card("p5_card_sos", 795, 145, 220, 82, 304, "Standards of Service %", "Standards"),
        card("p5_card_wait", 1025, 145, 225, 82, 305, "Average Wait Time", "Avg wait"),
        matrix(
            "p5_detail_matrix",
            35,
            245,
            1215,
            395,
            400,
            [
                ("Dim Date", "Date"),
                ("Dim Time Slot", "Half Hour"),
                ("Dim Team", "Team"),
                ("Dim Call Group", "Call Group"),
                ("Dim Queue", "Queue"),
                ("Dim Call Type", "Call Type"),
                ("Dim Emergency Classification", "Emergency Classification"),
                ("Dim Working Hours", "Working Hours"),
            ],
            [
                "Calls Offered",
                "Calls Answered",
                "Calls Abandoned",
                "Standards of Service %",
                "Average Speed To Answer",
                "Average Handling Time",
                "Total Talk Time",
            ],
            "Detailed aggregate records",
        ),
        textbox("p5_detail_note", 45, 665, 1160, 24, 401, "Detail rows reflect the source grain: date, half-hour and call type. Individual caller, agent and extension-level records are not available in the supplied extract.", 8, "#475569", False),
    ]
    create_page(page_ids[4], "Detailed Call Records", "Drill detail by date, time, mapped service line, outcome measures and duration metrics", page5)

    pages_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": page_ids,
        "activePageName": page_ids[0],
    }
    pages_json_path = PAGES_DIR / "pages.json"
    try:
        pages_json_path.write_text(json.dumps(pages_json, indent=2), encoding="utf-8")
    except PermissionError:
        pages_json_path.with_suffix(".json.generated").write_text(json.dumps(pages_json, indent=2), encoding="utf-8")


def write_model_notes(summary: dict[str, object]) -> None:
    content = f"""# Power BI Model Notes

## Source Profile

- Raw fact extract: `Source Data/raw/data.xlsx`, sheet `Data`.
- Call-type mapping: `Source Data/raw/mapping table.xlsx`, sheet `Mapping - Call Types`.
- Agent-team mapping: `Source Data/raw/mapping table.xlsx`, sheet `Mapping - Agent Team`.
- Source rows: {summary["source_rows"]:,}.
- Modelled fact rows: {summary["fact_rows"]:,}.
- Excluded rows: {summary["excluded_rows"]:,} rows have no Date and are kept in `Curated Data/excluded_no_date_rows.csv`.
- Date range: {summary["date_min"]} to {summary["date_max"]}.
- Call types mapped: {summary["call_type_count"]}.
- Teams mapped: {summary["team_count"]}.

## Modelling Assumptions

- The fact grain is date + half-hour + call type, not individual call records.
- `Calls Handled` is modelled as Calls Answered.
- `NG Calls Abandoned` and `Abandoned Within Service Level` are combined for the Calls Abandoned measure.
- `Service Level Calls / Service Level Calls Offered` is used for Standards of Service %.
- `Queue` is derived from Call Type because no separate queue field is present.
- Emergency classification is derived from mapped Emergency team membership, call types containing `Emergency`, and `False Alarm`.
- Duration fields are treated as seconds.
- `HandleTime` is used for handling time and as the closest available talk-time proxy.
- The target measure is set to 90% because no target table is provided.

## Unavailable From The Provided Source

- Individual agent or extension performance cannot be calculated from the fact extract.
- Calls transferred, hold time, after-call work time, logged-in time, repeat calls, caller identifiers and true call-level detail are not present.
- Placeholder measures for unavailable KPIs intentionally return blank.

## Remaining Manual Power BI Desktop Steps

- Open `Call Handling Standards of Service Dashboard.pbip` in Power BI Desktop and refresh.
- Review visual formatting after Desktop materialises the generated PBIR visuals.
- Add conditional formatting rules to matrices using `Below Service Target Flag`, `Standards of Service Variance to Target`, and `Abandonment Rate %`.
- If desired, configure the `Detailed Call Records` page as a formal drill-through page from queue, team and call-group visuals.
- Optionally adjust the `Service Level Target %` measure if the service standard differs from 90%.
"""
    (ROOT / "MODEL_NOTES.md").write_text(content, encoding="utf-8")


def main() -> None:
    summary = build_curated_data()
    build_semantic_model()
    build_report()
    write_model_notes(summary)
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()

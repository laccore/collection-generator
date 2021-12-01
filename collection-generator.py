import pandas as pd
import datetime
import simplekml
import timeit
import os.path
import sqlite3
from gooey import Gooey, GooeyParser

pd.options.mode.chained_assignment = None


def export_csv(dataframe, filename):
    export_start_time = timeit.default_timer()
    print("Exporting collection to CSV...", end="\r", flush=True)
    csv_data = dataframe[
        [
            "Location",
            "Country",
            "State_Province",
            "Hole_ID",
            "Original_ID",
            "Date",
            "Water_Depth",
            "Lat",
            "Long",
            "Elevation",
            "Position",
            "Sample_Type",
            "mblf_T",
            "mblf_B",
            "IGSN",
        ]
    ]
    csv_data["Date"] = csv_data["Date"].apply(
        lambda x: x.replace("-", "") if pd.notnull(x) else x
    )

    csv_data.to_csv(filename, encoding="utf-8-sig", index=False, float_format="%g")

    print(
        f"Collection exported to {filename} in {round(timeit.default_timer()-export_start_time,2)} seconds.",
        flush=True,
    )


def export_html(dataframe, filename):
    export_start_time = timeit.default_timer()
    print("Exporting collection to HTML...", end="\r", flush=True)
    html_data = (
        "<tr><td>"
        + dataframe["Location"].apply(lambda x: "" if pd.isnull(x) else str(x))
        + "</td><td>"
    )
    html_data += (
        dataframe["Lat"].apply(lambda x: "" if pd.isnull(x) else str(x)) + "</td><td>"
    )
    html_data += (
        dataframe["Long"].apply(lambda x: "" if pd.isnull(x) else str(x)) + "</td><td>"
    )
    html_data += (
        dataframe["Elevation"].apply(lambda x: "" if pd.isnull(x) else str(round(x)))
        + "</td><td>"
    )
    html_data += (
        dataframe["IGSN"].apply(lambda x: "" if pd.isnull(x) else x) + "</td></tr>"
    )

    with open(filename, "w") as f:
        for r in html_data:
            f.write(r + "\n")

    print(
        f"Collection exported to {filename} in {round(timeit.default_timer()-export_start_time,2)} seconds.",
        flush=True,
    )


def export_kml(dataframe, filename):
    export_start_time = timeit.default_timer()
    print("Exporting collection to KML...", end="\r", flush=True)

    kml_data = dataframe[["Location", "Long", "Lat"]]

    kml_data["GECF"] = (
        dataframe["Hole_ID"].apply(
            lambda x: "LacCoreID: " + str(x) if pd.notnull(x) else ""
        )
        + dataframe["Original_ID"].apply(
            lambda x: " / FieldID: " + str(x) if pd.notnull(x) else ""
        )
        + dataframe["Date"].apply(
            lambda x: " / Date: " + str(x) if pd.notnull(x) else ""
        )
        + dataframe["Water_Depth"].apply(
            lambda x: " / Water Depth: " + str(x) + "m " if pd.notnull(x) else ""
        )
        + dataframe[["mblf_T", "mblf_B"]].apply(
            lambda x: (
                " / Sediment Depth: "
                + (str(x[0]) if pd.notnull(x[0]) else "?")
                + "-"
                + (str(x[1]) if pd.notnull(x[1]) else "?")
                + "m"
            )
            if (pd.notnull(x[0]) or pd.notnull(x[1]))
            else "",
            axis=1,
        )
        + dataframe["Position"].apply(
            lambda x: " / Position: " + str(x) if pd.notnull(x) else ""
        )
        + dataframe["IGSN"].apply(
            lambda x: " / IGSN: " + str(x) if pd.notnull(x) else ""
        )
        + dataframe["Sample_Type"].apply(
            lambda x: " / Sample Type: " + str(x) if pd.notnull(x) else ""
        )
    )

    kml = simplekml.Kml(name="LacCore/CSDCO Core Collection")
    style = simplekml.Style()
    style.iconstyle.icon.href = (
        "http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png"
    )
    style.iconstyle.scale = 1.2
    style.iconstyle.color = "ff0000cc"

    for _, k in kml_data.iterrows():
        if pd.notnull(k[1]) and pd.notnull(k[2]):
            pnt = kml.newpoint(name=k[0], coords=[(k[1], k[2])], description=k[3])
            pnt.style = style

    kml.save(path=filename)

    print(
        f"Collection exported to {filename} in {round(timeit.default_timer()-export_start_time,2)} seconds.",
        flush=True,
    )


def process_holes(holes_database, filename):
    print(f"Loading {holes_database}...", end="\r", flush=True)
    load_start_time = timeit.default_timer()

    # Set up database connection in uri mode so as to open as read-only
    conn = sqlite3.connect("file:" + holes_database + "?mode=ro", uri=True)

    # Get all data from boreholes table
    dataframe = pd.read_sql_query("SELECT * FROM boreholes", conn)

    print(
        f"Loaded {holes_database} in {round(timeit.default_timer()-load_start_time,2)} seconds.",
        flush=True,
    )

    dataframe["Long"] = dataframe["Long"].apply(
        lambda x: None if pd.isnull(x) else round(x, 4)
    )
    dataframe["Lat"] = dataframe["Lat"].apply(
        lambda x: None if pd.isnull(x) else round(x, 4)
    )
    dataframe["Water_Depth"] = dataframe["Water_Depth"].apply(
        lambda x: None if pd.isnull(x) else round(x, 2)
    )

    print(flush=True)
    export_csv(dataframe=dataframe, filename=filename + ".csv")
    export_html(dataframe=dataframe, filename=filename + ".txt")
    export_kml(dataframe=dataframe, filename=filename + ".kml")
    print(flush=True)

    conn.close()

    print(
        f"Finished processing {holes_database} in {round(timeit.default_timer()-load_start_time,2)} seconds.",
        flush=True,
    )


@Gooey(program_name="CSDCO Collection Generator")
def main():
    default_path = ""
    potential_paths = [
        "/Volumes/CSDCO/Vault/projects/!inventory/CSDCO.sqlite3",
        "Z:/projects/!inventory/CSDCO.sqlite3",
        "Y:/projects/!inventory/CSDCO.sqlite3",
        "Z:/Vault/projects/!inventory/CSDCO.sqlite3",
        "Y:/Vault/projects/!inventory/CSDCO.sqlite3",
    ]

    for path in potential_paths:
        if os.path.isfile(path):
            default_path = path
            print(f"Found CSDCO database at '{path}'.")
            break

    parser = GooeyParser(
        description="Export borehole data from the CSDCO database for publishing."
    )

    input_output = parser.add_argument_group(
        "Input and Output", gooey_options={"columns": 1}
    )
    input_output.add_argument(
        "database_file",
        widget="FileChooser",
        metavar="CSDCO Database File",
        default=default_path,
        help="Path of the CSDCO database file.",
    )
    input_output.add_argument(
        "output_directory",
        widget="DirChooser",
        metavar="Save Path",
        help="Where to save output files.",
    )

    options = parser.add_argument_group("Export Options")
    options.add_argument(
        "-d",
        "--date-stamp",
        metavar="Append datestamp to file names",
        action="store_true",
        help="Export files with the date in the filename (e.g., collection_YYYYMMDD.csv).",
    )

    args = parser.parse_args()

    if not os.path.isfile(args.database_file):
        print(
            f"ERROR: database file '{args.database_file}' does not exist. Exiting.",
            flush=True,
        )
        exit(1)

    if not os.path.isdir(args.output_directory):
        print(
            f"ERROR: output folder '{args.output_directory}' does not exist. Exiting.",
            flush=True,
        )
        exit(1)

    export_filename = (
        "collection"
        if not args.date_stamp
        else "collection_" + datetime.datetime.now().strftime("%Y%m%d")
    )
    export_filename = os.path.join(args.output_directory, export_filename)
    process_holes(args.database_file, export_filename)


if __name__ == "__main__":
    main()

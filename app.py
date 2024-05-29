from shiny import App, render, ui, reactive
import pandas as pd

app_ui = ui.page_fluid(
    ui.panel_title("CSV File Upload and Statistics Display"),
    ui.layout_sidebar(
        ui.panel_sidebar(
            ui.input_file("file1", "Choose CSV File", accept=[".csv"]),
            ui.input_numeric("nRows", "Number of rows to display:", value=10, min=1),
            ui.output_ui("speciesSelect"),
            ui.download_button("downloadData", "Download Table as CSV")
        ),
        ui.panel_main(
            ui.h3("Uploaded Data"),
            ui.output_table("table"),
            ui.h3("Statistics"),
            ui.output_table("statsTable")
        )
    )
)

def server(input, output, session):
    @reactive.Calc
    def data():
        file_info = input.file1()
        if file_info is None or len(file_info) == 0:
            return None
        return pd.read_csv(file_info[0]["datapath"])
    
    @output
    @render.ui
    def speciesSelect():
        df = data()
        if df is None:
            return None
        species = ["All"] + df["Species"].unique().tolist()
        return ui.input_select("species", "Filter by Species:", choices=species, selected="All")
    
    @reactive.Calc
    def filteredData():
        df = data()
        if df is None:
            return None
        if input.species() == "All":
            return df
        return df[df["Species"] == input.species()]
    
    @output
    @render.table
    def table():
        df = filteredData()
        if df is None:
            return None
        return df.head(input.nRows())
    
    @reactive.Calc
    def stats():
        df = filteredData()
        if df is None:
            return None
        grouped = df.groupby("Species").agg(["mean", "min", "max", "count"])
        stats_df = grouped.stack(level=1).reset_index().rename(columns={"level_1": "Statistic"})
        return stats_df
    
    @output
    @render.table
    def statsTable():
        return stats()
    
    @session.download(filename=lambda: f"table-{pd.Timestamp('today').date()}.csv")
    def downloadData():
        df = filteredData()
        if df is not None:
            return df.to_csv(index=False)
        return ""

app = App(app_ui, server)

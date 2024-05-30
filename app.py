from shiny import App, render, ui, reactive
import pandas as pd
import plotly.express as px
from shinywidgets import output_widget, render_widget
from libreoffice_utils import load_extract

app_ui = ui.page_fluid(
    ui.panel_title("CSV File Upload and Statistics Display"),
    ui.layout_sidebar(
        ui.panel_sidebar(
            ui.input_file("file1", "Choose CSV File", accept=[".csv"]),
            ui.output_ui("speciesSelect"),
            ui.download_button("downloadData", "Download selected data as CSV")
        ),
        ui.panel_main(
            ui.h3("Box Plot for Petal Length"),
            ui.panel_well(output_widget("petalLengthPlot")),
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
    @render_widget
    def petalLengthPlot():
        df = filteredData()
        if df is None:
            return None
        fig = px.box(df, x="Species", y="Petal.Length", title="Petal Length by Species")
        return fig
    
    @reactive.Calc
    def stats():
        df = filteredData()
        if df is None:
            return None
        grouped = df.groupby("Species").agg(["mean", "min", "max"])
        stats_df = grouped.stack(level=1).reset_index().rename(columns={"level_1": "Statistic"})
        return stats_df
    
    @output
    @render.table
    def statsTable():
        return stats()
    
    @session.download(filename=lambda: f"table-{pd.Timestamp('today').date()}.csv")
    def downloadData():
        df = filteredData()
        df_out=load_extract(df)
        if df_out is not None:
            species=input.species()
            return df_out.to_csv(f'Statistics_{species}.csv',index=False)
        return ""

app = App(app_ui, server)

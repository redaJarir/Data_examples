from shiny import App, reactive, ui, render
import pandas as pd
import plotly.express as px
from shinywidgets import output_widget, render_widget
from libreoffice_utils import load_extract
from shinyswatch import theme

# Apply Bootswatch theme
theme.flatly()

# Inline CSS for custom styling
custom_css = """
<style>
    .custom-sidebar {
        background-color: #223f52 !important;
        padding: 10px;
        border-radius: 5px;
    }
    .custom-title {
        background-color: #122432 !important;
        color: white;
        padding: 10px;
        border-radius: 5px;
        text-align: center;
        font-weight: bold;
    }
    .btn {
        background-color: #8da5b8 !important;
        color: white !important;
        border: none !important;
    }
    .btn:hover {
        background-color: #122432 !important;
    }
    .custom-file-label, .custom-select-label {
        color: white !important;
    }
    h3 {
        color: #223f52 !important;
        font-weight: bold !important;
    }
    .custom-plot-container {
        background-color: #8da5b8 !important;
        border-radius: 5px;
        padding: 15px;
        box-shadow: 0px 0px 5px rgba(0, 0, 0, 0.1);
    }
</style>
"""

app_ui = ui.page_fluid(
    ui.HTML(custom_css),  # Include custom CSS
    ui.div(
        ui.h2("CSV File Upload and Statistics Display", class_="custom-title"), 
        class_="custom-title"
    ),
    ui.layout_sidebar(
        ui.panel_sidebar(
                        ui.div(
                ui.input_file("file1", "Choose CSV File", accept=[".csv"]),
                class_="custom-file-label"
            ),
            ui.output_ui("speciesSelect"),
            ui.download_button("downloadData", "Download selected data as CSV"),
            class_="custom-sidebar"
        ),
        ui.panel_main(
            ui.h3("Box Plot for Petal Length"),
            ui.panel_well(
                output_widget("petalLengthPlot"), 
                class_="custom-plot-container"
            ),
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
        return ui.div(
            ui.input_select("species", "Filter by Species:", choices=species, selected="All"),
            class_="custom-select-label"
        )
    
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
        fig.update_layout(
        #    paper_bgcolor='rgba(0,0,0,0)',
        #    plot_bgcolor='rgba(0,255,0,0.1)',  # Set your desired background color
            title_x=0.5  # Center the title
        )
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
        df_out = load_extract(df)
        df_out=load_extract(df)
        if df_out is not None:
            species=input.species()
            return df_out.to_csv(f'Statistics_{species}.csv',index=False)
        return ""

app = App(app_ui, server)

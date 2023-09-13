import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import xlsxwriter
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

def main():
    st.title("Company Metrics Visualization")

    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        # Get unique company names
        company_names = df['Company'].unique()

        # User selects a company
        selected_company = st.selectbox("Select a company", company_names)

        # Filter data for the selected company
        company_data = df[df['Company'] == selected_company]

        # Ask user to input metrics to plot
        st.sidebar.title("Select Metrics to Plot")
        num_metrics = st.sidebar.number_input("Number of metrics to plot", min_value=1, max_value=10, value=1, step=1)

        metrics_to_plot = []
        for i in range(num_metrics):
            st.sidebar.markdown(f"**Metric {i + 1}**")
            metric = st.sidebar.selectbox(f"Select metric {i + 1}", ['Sales', 'Profit', 'RoCE'])

            plot_type = st.sidebar.selectbox(f"Select plot type {i + 1}", ['Line', 'Bar'])

            y_axis = st.sidebar.selectbox(f"Select Y Axis {i + 1}", ['Left Y Axis', 'Right Y Axis'])

            metrics_to_plot.append({
                'metric': metric,
                'plot_type': plot_type,
                'y_axis': y_axis
            })

        # Create subplots for selected metrics
        fig, ax1 = plt.subplots(figsize=(10, 6))

        for metric_info in metrics_to_plot:
            metric = metric_info['metric']
            plot_type = metric_info['plot_type']
            y_axis = metric_info['y_axis']

            ax = ax1 if y_axis == 'Left Y Axis' else ax1.twinx()

            if plot_type == 'Line':
                ax.plot(company_data['Year'], company_data[metric], marker='o', label=metric)
            elif plot_type == 'Bar':
                ax.bar(company_data['Year'], company_data[metric], label=metric, alpha=0.5)

            ax.set_xlabel("Year")
            ax.set_ylabel(metric)

        fig.tight_layout()
        st.pyplot(fig)

        # Save the graph as an image
        image_path = "company_metrics_graph.png"
        fig.savefig(image_path, bbox_inches='tight', dpi=100)

        # Create a replicate Excel workbook and add the data and graph to it
        output = io.BytesIO()
        with xlsxwriter.Workbook(output, {'in_memory': True}) as workbook:
            # Add a worksheet and write the data
            worksheet = workbook.add_worksheet("Company Data")
            for i, col in enumerate(df.columns):
                worksheet.write(0, i, col)
            for i, value in enumerate(df.values):
                for j, val in enumerate(value):
                    worksheet.write(i + 1, j, val)

            # Insert the saved image at the end of the Excel file
            worksheet.insert_image('A1', image_path, {'x_offset': 0, 'y_offset': 0, 'x_scale': 0.5, 'y_scale': 0.5})

        # Allow the user to download the Excel file
        st.markdown("### Download Excel with Data and Graph")
        st.download_button(label="Download Excel", data=output.getvalue(), file_name="company_metrics.xlsx", key="download")

if __name__ == "__main__":
    main()

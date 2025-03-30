# Mekko Chart Generator for Excel

This VBA module provides a solution for generating Mekko (Marimekko) charts directly within Excel. It processes a user-selected data table, computes sub-segment totals, and renders a dynamic chart using Excel shapes. Compatible with both Windows and macOS versions of Excel. Key features include:

- **Data Processing:** Reads your selected table (including headers), calculates totals and sub-segment values, and sorts categories in descending order.
- **User Configuration:** Prompts for chart title, maximum number of categories (rows), and whether to display percentages.
- **Dynamic Chart Rendering:** Creates a new worksheet, draws the chart with color-coded sub-segments, builds a legend, and aggregates excess data into an "Other" category.
- **Custom Visuals:** Automatically extracts colors from the input data for both fills and fonts, ensuring consistency and visual appeal.

## Installation and Usage

1. **Import the Module:** Open the VBA editor in Excel:
   - Windows: Press `ALT + F11`
   - Mac: Press `FN + Option/Alt + F11` or `FN + ⌥ + F11`
   Then import this module into your workbook.
2. **Prepare Your Data:** Ensure your table is organized with headers in the first row and category labels in the first column.
3. **Select Your Data:** Highlight the entire table (headers included).
4. **Run the Macro:** Execute the `CreateMekkoChart` macro:
   - Windows: Press `ALT + F8`, select `CreateMekkoChart`, and click "Run"
   - Mac: Go to Tools > Macro > Macros or press `⌥ + F8`, select `CreateMekkoChart`, and click "Run"
5. **Follow Prompts:** Input the chart title, set the maximum number of categories, and decide if you want percentages displayed.
6. **View Your Chart:** A new worksheet with your generated Mekko chart will be created automatically.

## Contribution Guidelines

Contributions are welcome! To contribute:
- Fork the repository.
- Make improvements, bug fixes, or add new features.
- Submit a pull request with your changes.
- Open issues to report bugs or suggest enhancements.

Please maintain consistent coding style and include comments for clarity.

## License

This project is licensed under the **MIT License**. This permissive license encourages collaboration by allowing anyone to use, modify, and distribute the code with minimal restrictions. See the `LICENSE` file for full details.

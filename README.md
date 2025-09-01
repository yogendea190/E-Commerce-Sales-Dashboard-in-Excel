# E-Commerce Sales Dashboard in Excel

## Project Overview

This project develops an interactive **E-Commerce Sales Dashboard** using Microsoft Excel to analyze sales and profit trends for an online retail company. The dashboard leverages advanced Excel features like pivot tables, charts, slicers, and combo boxes to provide dynamic insights into sales performance across product categories, regions, and months. The project demonstrates skills in **data analysis**, **data visualization**, and **interactive reporting**, making it ideal for MIS Executive roles.

**Domain**: E-Commerce  
**Tools Used**: Microsoft Excel (Pivot Tables, Charts, Slicers, Combo Boxes, Data Analysis Toolpak)  
**GitHub Repository**: [Link to your GitHub repo]

---

## Objective

The goal is to create a user-friendly dashboard that allows stakeholders to:
- Analyze sales and profit trends by product category, month, and region.
- Use a combo box for interactive filtering of product categories.
- Visualize shipping days through a histogram.
- Provide actionable insights for data-driven decision-making.

---

## Dataset Description

The dataset (`E-Commerce Dashboard.xlsx`) contains order data for an e-commerce company with the following fields:

| **Field**          | **Description**                                      |
|--------------------|------------------------------------------------------|
| Order ID           | Unique Order ID of a product                          |
| Order Date         | Order Placement Date                                 |
| Ship Date          | Shipment Date of the placed order                    |
| Aging              | Used to create histogram bins for shipping days      |
| Ship Mode          | Shipment mode of placed order                        |
| Product Category   | Product Category (e.g., Electronics, Clothing)        |
| Product            | Name of the Product                                  |
| Sales              | Sales Amount                                         |
| Quantity           | Number of items ordered                              |
| Discount           | Deduction from the usual cost                        |
| Profit             | Financial advantage or benefit                       |
| Shipping Cost      | Amount required to ship the order                    |
| Order Priority     | Precedence of placed order                           |
| Customer ID        | Unique Customer ID                                   |
| Customer Name      | Name of the Customer                                 |
| Segment            | Product Segment (e.g., Home Office, Corporate)       |
| City               | Unique City Name                                     |
| State              | Unique State Name                                    |
| Country            | Unique Country Name                                  |
| Region             | Part of a country                                    |
| Months             | Month of placing the order                           |

**Dataset File**: [E-Commerce Dashboard.xlsx](dataset/E-Commerce_Dashboard.xlsx)

---

## Analysis Tasks

The project involves the following tasks:
1. **Create a Histogram**: Analyze the number of shipping days (Aging).
2. **Prepare Tables**:
   - Month-wise Sales and Profit table.
   - Region-wise Sales table.
3. **Create a Combo Box**: Add user control for filtering by Product Category.
4. **Create Charts**:
   - Column chart for month-wise Sales and Profit.
   - Column chart for region-wise Sales.
5. **Link Tables with Combo Box**: Enable dynamic filtering of data.
6. **Build a Dashboard**: Combine all visuals and controls for interactive reporting.

---

## Step-by-Step Implementation

### 1. Create a Histogram for Shipping Days (Aging)
- **Objective**: Visualize the distribution of shipping days.
- **Steps**:
  1. Open the dataset (`E-Commerce Dashboard.xlsx`).
  2. Go to **Data Tab** > **Analysis Group** > **Data Analysis** > **Histogram**.
  3. In the Histogram dialog box:
     - Check **Labels** (since data includes headers).
     - **Input Range**: Select `Sales Data!D1:D51291` (Aging column).
     - **Bin Range**: Select `Working!K3:K7` (predefined bins for shipping days).
     - **Output Range**: Select `Working!N3` for the histogram table.
     - Check **Chart Output** to generate the histogram.

### 2. Prepare Tables in the Working Sheet
- **Month-wise Sales and Profit Table**:
  1. Create a pivot table in the `Working` sheet.
  2. Set **Rows** as `Months`, **Values** as `Sum of Sales` and `Sum of Profit`.
  3. Format the table for clarity (e.g., currency format for Sales and Profit).
- **Region-wise Sales Table**:
  1. Create another pivot table in the `Working` sheet.
  2. Set **Rows** as `Region`, **Values** as `Sum of Sales`.
  3. Apply conditional formatting to highlight key regions.

### 3. Create a User Control Combo Box
- **Objective**: Allow users to filter data by Product Category.
- **Steps**:
  1. Go to **Developer Tab** > **Insert** > **Combo Box (Form Control)**.
  2. Place the combo box in the dashboard sheet.
  3. Link it to a cell (e.g., `Dashboard!A1`) for output.
  4. Populate the combo box with `Product Category` values using a range (e.g., `Sales Data!F1:F51291`).
  5. Use **VLOOKUP** or **INDEX-MATCH** to filter data based on the selected category.

### 4. Create Column Charts
- **Month-wise Chart**:
  1. Select the month-wise Sales and Profit table.
  2. Insert a **Column Chart** (stacked or clustered).
  3. Customize with titles, labels, and colors for clarity.
- **Region-wise Chart**:
  1. Select the region-wise Sales table.
  2. Insert a **Column Chart**.
  3. Format to highlight regional performance.

### 5. Link Tables with Combo Box
- Use Excel formulas (e.g., **IF**, **VLOOKUP**, or **Pivot Table Filters**) to dynamically update tables and charts based on the selected Product Category in the combo box.
- Add **Slicers** for additional interactivity (e.g., filter by Region or Ship Mode).

### 6. Create the Dashboard
- **Steps**:
  1. Create a new sheet named `Dashboard`.
  2. Arrange the combo box, slicers, charts, and tables in a visually appealing layout.
  3. Add titles, logos, and color schemes to enhance professionalism.
  4. Ensure all elements are linked for interactivity.

---


# Excel Scatterplot Dash App

This simple Dash app allows users to upload an Excel file (.xlsx) with at least 3 columns.
The first two columns are used as the x and y axes, and the third as the color group in a scatter plot.

##
2 axis scatter
one-to-5 scale on both axis
Axis Y: Impact materiality (see column "Impact")
Axis X: Financial materiality (risk or opportunity)  -  (see column "Risk")

Hover text for each dot in defined by column "Name of IRO"
Color is defined by column "Sub-Topic"

## Features

- Drag and drop file upload
- Interactive Plotly scatter plot
- Clean, white minimal layout

## Installation

```bash
pip install -r requirements.txt
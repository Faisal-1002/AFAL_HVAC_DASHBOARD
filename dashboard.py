import dash
from dash import dcc, html, dash_table  # Import dash_table from dash
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
from dash.dependencies import Input, Output
from datetime import time

# Load and clean the data from the local Excel file
file_path = "August Daily Reports.xlsx"  # Your local file path
df = pd.read_excel(file_path, engine='openpyxl')

# Clean the data
df_cleaned = df.drop_duplicates()
df_cleaned = df_cleaned.dropna(how='all')

# Function to convert time to hours
def time_to_hours(t):
    if isinstance(t, time):  # Check if it's a datetime.time object
        return t.hour + t.minute / 60 + t.second / 3600  # Convert to hours
    return 0  # Return 0 if the value is not valid

# Ensure 'Total Duration' is in hours and numeric format
df_cleaned['Total Duration'] = df_cleaned['Total Duration'].apply(time_to_hours)

# Drop rows where 'Total Duration' is NaN or remains non-numeric after conversion
df_cleaned = df_cleaned.dropna(subset=['Total Duration'])

# Convert 'Date' to a proper datetime format for filtering
df_cleaned['Date'] = pd.to_datetime(df_cleaned['Date'], errors='coerce')
df_cleaned = df_cleaned.dropna(subset=['Date'])  # Drop rows where Date is NaN

# Filter only finished jobs
df_cleaned = df_cleaned[df_cleaned['Job Status'] == 'Finished']

# Function to convert numeric hours to hh:mm format for the y-axis labels and hover
def hours_to_hhmm(hours):
    h = int(hours)
    m = int((hours - h) * 60)
    return f"{h:02d}:{m:02d}"

# Function to count occurrences of "Yes" in the columns
def count_yes(value):
    return 1 if value == 'Yes' else 0

# Function to create the scatter plot with lines, tolerance range, and mean line
def plot_correlation_graph(filtered_df):
    # Ensure 'Total Duration' is in numeric form for further calculations
    filtered_df['Total Duration'] = pd.to_numeric(filtered_df['Total Duration'], errors='coerce')

    # Mean value for Total Duration
    mean_value = filtered_df['Total Duration'].mean()

    # Create the figure object
    fig = go.Figure()

    # Add the line connecting the dots with markers
    fig.add_trace(go.Scatter(
        x=list(range(len(filtered_df))),  # Equal spacing for dots
        y=filtered_df['Total Duration'],
        mode='lines+markers',
        name='Total Duration',
        line=dict(color='blue'),
        marker=dict(size=8, color='blue'),
        hovertemplate=
        '<b>Activity Number</b>: %{text}<br>' +
        '<b>Total Duration</b>: %{customdata}<br>' +
        '<extra></extra>',
        text=filtered_df['Activity Number'],
        customdata=[hours_to_hhmm(y) for y in filtered_df['Total Duration']]  # Only show time format
    ))

    # Mark outliers with black question marks, placed next to the dot
    outlier_activity_numbers = []
    for i, value in enumerate(filtered_df['Total Duration']):
        if value > 5 or value < 0:  # Outliers outside the tolerance range
            outlier_activity_numbers.append(filtered_df['Activity Number'].iloc[i])
            fig.add_annotation(
                x=i + 0.1,  # Shift question mark slightly to the right
                y=value + 0.2,  # Shift question mark slightly upwards
                text='?',
                showarrow=False,
                font=dict(size=16, color='black')
            )

    # Add the upper tolerance range line (5 hours)
    fig.add_hline(
        y=5,
        line=dict(color='orange', width=2, dash='dash'),
        annotation_text="Upper Tolerance (5 hours)",
        annotation_position="top left"
    )

    # Add the lower tolerance range line (0 hours)
    fig.add_hline(
        y=0,
        line=dict(color='orange', width=2, dash='dash'),
        annotation_text="Lower Tolerance (0 hours)",
        annotation_position="bottom left"
    )

    # Add the mean line
    fig.add_hline(
        y=mean_value,
        line=dict(color='green', width=2, dash='dash'),
        annotation_text=f"Mean = {hours_to_hhmm(mean_value)} hours",
        annotation_position="bottom right"
    )

    # Update y-axis ticks to show whole hours (0, 1, 2, etc.) and format hover with time
    y_ticks = list(range(0, int(filtered_df['Total Duration'].max()) + 2))
    y_labels = [hours_to_hhmm(y) for y in y_ticks]

    # Update the layout to include extra space on the y-axis and configure the axis labels
    fig.update_layout(
        title="Total Duration vs Activity Number",
        xaxis=dict(
            tickmode='array',
            tickvals=list(range(len(filtered_df))),  # Equal spacing for the dots
            ticktext=filtered_df['Activity Number'].astype(str),  # Show Activity Number as labels
            title='Activity Number'
        ),
        yaxis=dict(
            tickmode='array',
            tickvals=y_ticks,
            ticktext=y_labels,
            title='Total Duration (hh:mm)',
            range=[-1, max(y_ticks) + 1],  # Extend the y-axis range for more space
            dtick=1  # Set y-ticks to whole numbers
        ),
        height=800,  # Restore original height
        margin=dict(l=150, r=40, b=160, t=80),  # Adjust margins for more spacing between graph and table
        font=dict(size=12),  # Restore original font size
        showlegend=False
    )

    return fig, outlier_activity_numbers

# Function to create the table of outliers
def create_outliers_table(outlier_activity_numbers):
    data = [{'Activity Number': num, 'Justification': ''} for num in outlier_activity_numbers]
    
    return dash_table.DataTable(
        columns=[
            {'name': 'Activity Number', 'id': 'Activity Number'},
            {'name': 'Justification', 'id': 'Justification'}
        ],
        data=data,
        style_table={'width': '80%', 'margin': '0 auto 60px auto'},  # Add more bottom margin to space table from graph
        style_cell={'textAlign': 'center', 'padding': '10px'},  # Restore padding and center text
        style_header={'backgroundColor': 'lightgray', 'fontWeight': 'bold'},
        editable=True  # Allow the Justification column to be editable
    )

# Calculate additional metrics for the second row of boxes
def calculate_metrics(filtered_df):
    # Convert 'Response Time' and 'Repair Time' to numeric (hours) before calculating the mean
    if 'Response Time' in filtered_df.columns:
        filtered_df['Response Time (hours)'] = filtered_df['Response Time'].apply(time_to_hours)
        avg_response_time = filtered_df['Response Time (hours)'].mean()
    else:
        avg_response_time = 0

    if 'Repair Time' in filtered_df.columns:
        filtered_df['Repair Time (hours)'] = filtered_df['Repair Time'].apply(time_to_hours)
        avg_repair_time = filtered_df['Repair Time (hours)'].mean()
    else:
        avg_repair_time = 0

    # Count occurrences of "Yes" in Late Response and Late Repair columns
    num_late_responses = filtered_df['Late Response'].apply(count_yes).sum() if 'Late Response' in filtered_df.columns else 0
    num_late_repairs = filtered_df['Late Repair'].apply(count_yes).sum() if 'Late Repair' in filtered_df.columns else 0

    return {
        'avg_response_time': hours_to_hhmm(avg_response_time),  # Convert to hh:mm for display
        'avg_repair_time': hours_to_hhmm(avg_repair_time),  # Convert to hh:mm for display
        'num_late_responses': int(num_late_responses),
        'num_late_repairs': int(num_late_repairs)
    }

# Initialize the Dash app with Bootstrap theme
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# App layout with dropdowns, two rows of boxes, graph, and table
app.layout = dbc.Container(
    [
        dbc.Row(
            dbc.Col(html.H1("Ticket Dashboard", className="text-center my-4", style={'font-size': '24px'}), width=12)  # Restore original title size
        ),
        dbc.Row(
            [
                dbc.Col(
                    dcc.Dropdown(
                        id='date-dropdown',
                        options=[{'label': str(date), 'value': str(date)} for date in df_cleaned['Date'].dt.date.unique()],
                        placeholder="Select a date",
                        style={"width": "100%", "margin-bottom": "20px", 'font-size': '12px'}
                    ),
                    width=6
                ),
                dbc.Col(
                    dcc.Dropdown(
                        id='zone-dropdown',
                        options=[{'label': zone, 'value': zone} for zone in df_cleaned['Zone'].unique()],
                        placeholder="Select a zone",
                        style={"width": "100%", "margin-bottom": "20px", 'font-size': '12px'}
                    ),
                    width=6
                ),
            ],
            className="mb-4"
        ),
        # First row of boxes
        dbc.Row(
            [
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Total Tickets", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='total-tickets', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="primary",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Closed Tickets", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='closed-tickets', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="success",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Pending Tickets", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='pending-tickets', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="warning",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Repeated Calls", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='repeated-calls', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="danger",
                        inverse=True,
                    ),
                    width=3,
                ),
            ],
            className="mb-4",
        ),
        # Second row of boxes (New row for average response/repair times, late responses/repairs)
        dbc.Row(
            [
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Average Response Time", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='avg-response-time', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="info",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Average Repair Time", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='avg-repair-time', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="secondary",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Late Responses", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='num-late-responses', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="warning",
                        inverse=True,
                    ),
                    width=3,
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Late Repairs", style={'textAlign': 'center', 'font-size': '14px'}),
                            dbc.CardBody(html.H4(id='num-late-repairs', className="card-text", style={'textAlign': 'center', 'font-size': '14px'})),
                        ],
                        color="danger",
                        inverse=True,
                    ),
                    width=3,
                ),
            ],
            className="mb-4",
        ),
        # Graph row
        dbc.Row(
            dbc.Col(
                dcc.Graph(id='correlation-graph', style={'height': '600px'}),  # Restore original graph size
                width=12
            )
        ),
        # Table row for outliers with extra spacing
        dbc.Row(
            dbc.Col(
                html.Div(id='outliers-table', style={'marginTop': '100px'})  # Added larger margin for more space between graph and table
            )
        )
    ],
    fluid=True,
)

# Callback to update the graph, metrics, and statistics based on selected filters
@app.callback(
    [Output('correlation-graph', 'figure'),
     Output('total-tickets', 'children'),
     Output('closed-tickets', 'children'),
     Output('pending-tickets', 'children'),
     Output('repeated-calls', 'children'),
     Output('avg-response-time', 'children'),
     Output('avg-repair-time', 'children'),
     Output('num-late-responses', 'children'),
     Output('num-late-repairs', 'children'),
     Output('outliers-table', 'children')],
    [Input('date-dropdown', 'value'),
     Input('zone-dropdown', 'value')]
)
def update_dashboard(selected_date, selected_zone):
    filtered_df = df_cleaned

    # Filter by selected date
    if selected_date is not None:
        selected_date = pd.to_datetime(selected_date).date()  # Convert string to date
        filtered_df = filtered_df[filtered_df['Date'].dt.date == selected_date]

    # Filter by selected zone
    if selected_zone is not None:
        filtered_df = filtered_df[filtered_df['Zone'] == selected_zone]

    # Calculate stats for the dashboard
    total_tickets = filtered_df['Activity Number'].count()
    closed_tickets = filtered_df.shape[0]
    pending_tickets = 0  # Since we're only showing finished jobs, this will always be 0
    repeated_calls = filtered_df[filtered_df['Repeated Call'] == 'Yes'].shape[0]

    # Calculate additional metrics for the second row of boxes
    metrics = calculate_metrics(filtered_df)

    # Update the correlation graph and capture the outliers
    graph, outlier_activity_numbers = plot_correlation_graph(filtered_df)

    # Create the outliers table
    outliers_table = create_outliers_table(outlier_activity_numbers)

    return (graph, total_tickets, closed_tickets, pending_tickets, repeated_calls,
            metrics['avg_response_time'], metrics['avg_repair_time'], metrics['num_late_responses'],
            metrics['num_late_repairs'], outliers_table)

# Run the app
if __name__ == "__main__":
    app.run_server(debug=True)

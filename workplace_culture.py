import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from io import BytesIO
import textwrap

def create_culture_report(subgroup, subgroup_value, data):
    # Convert the data to a DataFrame if it's not already
    data = pd.DataFrame(data)

    # Filter the data for the specified subgroup
    filtered_data = data[data[subgroup] == subgroup_value]

    # Extract relevant columns
    columns_of_interest = [
        "Our leaders treat staff with respect.",
        "Staff treat each other with respect.",
        "My co-workers trust in me and each other.",
        "In my area, employees are treated fairly.",
        "I am empowered to investigate problems and explore new ideas at work."
    ]

    survey_data = filtered_data[columns_of_interest]

    # Calculate percentages for each response type
    response_types = ["Strongly disagree", "Disagree", "Neutral", "Agree", "Strongly agree"]
    percentages = {col: survey_data[col].value_counts(normalize=True).reindex(response_types, fill_value=0) * 100 for col in columns_of_interest}

    data_values = np.array([
        percentages[columns_of_interest[0]].values,
        percentages[columns_of_interest[1]].values,
        percentages[columns_of_interest[2]].values,
        percentages[columns_of_interest[3]].values,
        percentages[columns_of_interest[4]].values
    ])

    # Colors for each category
    colors = ['#F24837', '#FC8F3E', '#FCCE48', '#C8DA51', '#5FBE78']
    categories = ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree']

    fig, ax = plt.subplots(figsize=(14, 8))  # Increase figure width for wider labels

    # Plotting the data
    cumulative_data = np.cumsum(data_values, axis=1)
    for i, (color, category) in enumerate(zip(colors, categories)):
        ax.barh(columns_of_interest, data_values[:, i], left=cumulative_data[:, i] - data_values[:, i], color=color, label=category, edgecolor='white', height=0.5)  # Increase bar height

    # Adding percentage labels within the blocks
    for j in range(data_values.shape[0]):
        for i in range(data_values.shape[1]):
            if data_values[j, i] >= 5:
                ax.text(cumulative_data[j, i] - data_values[j, i] / 2, j, f"{data_values[j, i]:.0f}%", va='center', ha='center', color='white', fontsize=16, fontweight='bold')

    # Wrap y-tick labels and hide y-ticks
    wrap_width = 30  # Set your desired wrap width here
    ax.set_yticklabels([textwrap.fill(label, wrap_width) for label in columns_of_interest], fontsize=14, weight='bold')
    ax.yaxis.set_ticks_position('none')

    # Customizing the plot to match the desired style
    ax.legend(bbox_to_anchor=(0.5, 0), loc='upper center', ncol=5, frameon=False, prop={'size': 12, 'weight': 'bold'})
    ax.set_xlim(0, 100)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)

    plt.tight_layout()
    plt.subplots_adjust(left=0.3)  # Adjust the left margin to provide more space for y-axis labels

    # Save the plot to a BytesIO object
    buf = BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)

    return buf

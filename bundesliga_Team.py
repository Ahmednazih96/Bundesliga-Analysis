import pandas as pd
import plotly.express as px
import time
# C:\Users\user\OneDrive\Data Curation Progect File\bundesliga_Team.py

# file_path = r"C:\\Users\\user\\Downloads\\bundesliga23_24\\accurate_cross_team.csv"

# df = pd.read_csv(file_path, encoding='ISO-8859-1')  

#folder_path = r"C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24" 


file_path = file_path = r'C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24\bundesliga_Team.xlsx'

all_sheets = pd.read_excel(file_path, sheet_name=None)

print("Sheets Names:", all_sheets.keys())

for sheet_name, df in all_sheets.items():
    print(f"\nData from sheet: {sheet_name}")
    print(df.head())  
# output_excel = r"C:\Users\user\OneDrive\Data Curation Progect File\Team\bundesliga_Team.xlsx"  







# with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
#     for filename in os.listdir(folder_path):
#         if filename.endswith('.csv'):  
#             file_path = os.path.join(folder_path, filename)
#             sheet_name = os.path.splitext(filename)[0] 
#             df = pd.read_csv(file_path)  
#             df.to_excel(writer, index=False, sheet_name=sheet_name)  





# team_name_corrections = {
#         'Bayern MÃ¼nchen': 'Bayern München',
#     'Borussia MÃ¶nchengladbach': 'Borussia Mönchengladbach',
#     'FC KÃ¶ln': 'FC Köln',
    
# }

# df['name'] = df['name'].apply(lambda x: team_name_corrections.get(x, x))

# try:
#     df.to_csv(file_path,index=False, encoding='ISO-8859-1')  # أو 'utf-16'
#     print("File saved successfully")
#     print(df.head())
# except Exception as e:
#     print(f"Error saving the file: {e}")

#1. Team Performance Comparison (Home vs. Away vs. Overall)
# "C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24\Bundesliga_table_2023_24.xlsx"

home_table_df = pd.read_csv(r"C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24\Bundesliga_table_home_2023_24.csv")
away_table_df = pd.read_csv(r"C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24\Bundesliga_table_away_2023_24.csv")
overall_table_df = pd.read_csv(r"C:\Users\user\OneDrive\Data Curation Progect File\bundesliga23_24\Bundesliga_table_2023_24.csv")

home_table_df['location'] = 'Home'
away_table_df['location'] = 'Away'
overall_table_df['location'] = 'Overall'

columns_to_select = ['name', 'pts', 'wins', 'draws', 'losses', 'goalConDiff', 'location']
combined_df = pd.concat([
    home_table_df[columns_to_select],
    away_table_df[columns_to_select],
    overall_table_df[columns_to_select]
], ignore_index=True)

fig = px.bar(
    combined_df, 
    x='name', 
    y='pts', 
    color='location', 
    title='Comparison of Points: Home vs Away vs Overall', 
    labels={'pts': 'Points', 'name': 'Team'},
    category_orders={'location': ['Home', 'Away', 'Overall']}
)

fig.update_layout(xaxis_tickangle=-90)
fig.show(renderer='iframe_connected')
fig.write_html("Comparison of Points.html")
time.sleep(1)

#2. Home Advantage Analysis (Points Difference)

home_vs_away = home_table_df[['name', 'pts']].set_index('name').join(away_table_df[['name', 'pts']].set_index('name'), rsuffix='_away')
home_vs_away['pts_diff'] = home_vs_away['pts'] - home_vs_away['pts_away']

fig = px.bar(home_vs_away, x=home_vs_away.index, y='pts_diff', 
             title='Home vs Away Points Difference', 
             labels={'pts_diff': 'Points Difference', 'index': 'Team'})

fig.update_layout(xaxis_tickangle=-90)
fig.show(renderer='iframe_connected')
fig.write_html("Home vs Away Points Difference.html")
time.sleep(1)





#3. Top Performing Teams (Overall Performance)

top_overall_teams = overall_table_df.sort_values(by='pts', ascending=False).head(10)

# Plot the top 10 teams by points
fig = px.bar(top_overall_teams, x='name', y='pts', 
             title='Top 10 Teams by Total Points', 
             labels={'pts': 'Points', 'name': 'Team'})

fig.update_layout(xaxis_tickangle=-90)
fig.show(renderer='iframe_connected')
fig.write_html("Top 10 Teams by Total Points.html")
time.sleep(1)


#4. Goal Scoring and Conceding Analysis

home_table_df['goals_scored'] = home_table_df['scoresStr'].apply(
    lambda x: int(x.split('-')[0]) if '-' in x and x.split('-')[0].isdigit() else 0
)
home_table_df['goals_conceded'] = home_table_df['scoresStr'].apply(
    lambda x: int(x.split('-')[1]) if '-' in x and x.split('-')[1].isdigit() else 0
)

away_table_df['goals_scored'] = away_table_df['scoresStr'].apply(
    lambda x: int(x.split('-')[0]) if '-' in x and x.split('-')[0].isdigit() else 0
)
away_table_df['goals_conceded'] = away_table_df['scoresStr'].apply(
    lambda x: int(x.split('-')[1]) if '-' in x and x.split('-')[1].isdigit() else 0
)


# Combine for comparison
combined_goals = pd.concat([home_table_df[['name', 'goals_scored', 'goals_conceded', 'location']],
                            away_table_df[['name', 'goals_scored', 'goals_conceded', 'location']]])

# Plot goals scored vs goals conceded
fig = px.bar(combined_goals, x='name', y=['goals_scored', 'goals_conceded'], color='location', 
             title='Goals Scored vs Goals Conceded: Home vs Away', 
             labels={'goals_scored': 'Goals Scored', 'goals_conceded': 'Goals Conceded', 'name': 'Team'})

fig.update_layout(barmode='group', xaxis_tickangle=-90)
fig.show(renderer='iframe_connected')
fig.write_html("Goals Scored vs Goals Conceded.html")
time.sleep(1)

# print(home_table_df[home_table_df['scoresStr'].str.contains(r'\D', na=False)]['scoresStr'])
# print(away_table_df[away_table_df['scoresStr'].str.contains(r'\D', na=False)]['scoresStr'])


#5. Team Consistency Analysis (Win Percentage)


home_table_df['win_percentage'] = home_table_df['wins'] / home_table_df['played'] * 100
away_table_df['win_percentage'] = away_table_df['wins'] / away_table_df['played'] * 100
overall_table_df['win_percentage'] = overall_table_df['wins'] / overall_table_df['played'] * 100

# Plot win percentage for home teams
fig = px.bar(home_table_df.sort_values(by='win_percentage', ascending=False).head(10),
             x='name', y='win_percentage', 
             title='Top 10 Teams by Win Percentage (Home)', 
             labels={'win_percentage': 'Win Percentage', 'name': 'Team'})

fig.update_layout(xaxis_tickangle=-90)
fig.show(renderer='iframe_connected')
fig.write_html("Top 10 Teams by Win Percentage (Home).html")
time.sleep(1)


#6. Points Distribution Analysis


combined_points = pd.concat([home_table_df[['name', 'pts']],
                             away_table_df[['name', 'pts']],
                             overall_table_df[['name', 'pts']]])

# Plot points distribution
fig = px.histogram(combined_points, x='pts', nbins=20, 
                   title='Distribution of Points Across Teams', 
                   labels={'pts': 'Points'})

fig.update_layout(bargap=0.1)
fig.write_html("Distribution of Points Across Teams.html")
time.sleep(1)






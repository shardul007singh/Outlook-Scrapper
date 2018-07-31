# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 11:45:29 2018

@author: shardsin
"""
import pandas as pd
import plotly.figure_factory as ff
from plotly.offline import plot
import re

dataFrame = pd.read_csv("Mails.csv")
userName = input("Kindly enter the name of the Engineer you want to search: ")

if re.search(userName, str(dataFrame['Sender']), re.IGNORECASE):
    newTaskTime = (
            dataFrame.loc[dataFrame['Subject'].str.contains("new task")]
            )['Received Time'].values
    
    taskStartedTime = (
            dataFrame.loc[dataFrame['Subject'].str.contains("started work")]
            )['Received Time'].values
    
    taskCompletedTime = (
            dataFrame.loc[dataFrame['Subject'].str.contains("completed")]
            )['Received Time'].values

    diff = pd.to_datetime(taskCompletedTime)[0] - pd.to_datetime(newTaskTime)[0]
    print("Total time taken to complete the task = {}".format(diff))

df = [
      dict(Task='Adhoc 1', 
          Start=str(pd.to_datetime(newTaskTime)[0]), 
          Finish=str(pd.to_datetime(taskStartedTime)[0]), 
          Resource='New'),
    dict(Task='Adhoc 1', 
         Start=str(pd.to_datetime(taskStartedTime)[0]), 
         Finish=str(pd.to_datetime(taskCompletedTime)[0]), 
         Resource='Progress'),
    dict(Task='Adhoc 1', 
         Start=str(pd.to_datetime(taskCompletedTime)[0]), 
         Finish=str(pd.to_datetime(taskCompletedTime)[0]), 
         Resource='Completed')
    ]

colors = dict(New = 'rgb(220, 0, 0)',
              Progress = (1, 0.9, 0.16),
              Completed = 'rgb(0, 220, 0)')
fig = ff.create_gantt(df, colors=colors, index_col='Resource', 
                      show_colorbar=True, showgrid_x=True, showgrid_y=True,
                      group_tasks=True,
                      title='ADHOC ACTIVITIES',
                      bar_width=0.02)
fig['data'][0]['marker'] = {'color': 'rgb(220, 0, 0)'}
fig['data'][0]['name'] = 'New Task'
fig['data'][1]['marker'] = {'color': (1, 0.9, 0.16)}
fig['data'][1]['name'] = 'Progress Started'
fig['data'][2]['marker'] = {'color': 'rgb(0, 220, 0)'}
fig['data'][2]['name'] = 'Task Completed'
fig['data'][3]['marker'] = {'color': 'rgb(0, 220, 0)', 'size': 10}
fig['layout']['showlegend'] = False
plot(fig, filename = 'Adhoc Task')




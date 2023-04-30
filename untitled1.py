import PySimpleGUI as sg
import docx2pdf
import webbrowser
from docx import Document
import pickle
import numpy as np
import helper
from datetime import date


col=helper.col2()
pickled_model = pickle.load(open('model1.pkl', 'rb'))

# Define the layout of the input form

layout = [
    [sg.Column([[sg.Text('Name:'), sg.InputText(key='name', size=(20, 1))],
                [sg.Text('Gender:'), sg.InputCombo(['Male', 'Female'], size=(17, 1), key='gender')],
                [sg.Text('Age :'), sg.InputText(key='age', size=(21, 1))],
                [sg.Text('Contact No:'), sg.InputText(key='contact', size=(16, 1))],
                
                [sg.Text('BP:'), sg.InputText(key='bp', size=(22, 1))],
                [sg.Text('Temp :'), sg.InputText(key='temp', size=(20, 1))],
                [sg.Text('SpO2 :'), sg.InputText(key='spo2', size=(20, 1))],
                
    ]),
     sg.Column([[sg.Text(col[i]), sg.Radio("Yes", group_id=i, key=i), sg.Radio("No", group_id=i,default=True, key=i)] for i in range(0,14)], element_justification='l'),
     sg.Column([[sg.Text(col[i]), sg.Radio("Yes", group_id=i, key=i), sg.Radio("No", group_id=i,default=True, key=i)] for i in range(14,28)], element_justification='l'),
     sg.Column([[sg.Text(col[i]), sg.Radio("Yes", group_id=i, key=i), sg.Radio("No", group_id=i,default=True, key=i)] for i in range(28,42)], element_justification='l'),
     
    ]
]

layout.append([sg.Button('Submit')],)
message=''
layout.append([sg.Text('Result:', size=(20, 1)), sg.Text(message, size=(20, 1), key='result')])
# Create the PySimpleGUI window with the layout
window = sg.Window('User Input Form', layout)

# Define the path of the docx file to be opened and updated
docx_path = "C:\\Users\\Hp\\OneDrive\\Sem6\\project\\R_copy.docx"

# Loop to keep the window open until user closes or submits the form
while True:
    event, values = window.read()
    
    # If user clicks cancel or closes the window, exit the loop
    if event == sg.WINDOW_CLOSED:
        break

    # If user clicks submit, display the input values and result in the output text field, update the docx file, and open the pdf in a browser
    if event == 'Submit':
        x_find=[]
        for i in range(0,42):
            
            if values[i]==True:
                x_find.append(1)
            else:
                    x_find.append(0)
        
    
        
        y = np.array(x_find)
        y
        y_find=pickled_model.predict(y.reshape(1, -1))
        
        val=str(y_find)
        message=y_find[0]
        
        # Open the docx file and update the values of name, gender, and temperature
        doc = Document(docx_path)
        today = date.today()
        def dsa(a,b):
            if a in paragraph.text:
                paragraph.text = paragraph.text.replace(a,b)
                    
        for paragraph in doc.paragraphs:
            dsa('P_name ',values['name'])
            dsa('Possible_sym',str(helper.get_symptoms(y_find[0])))
            dsa('2023-04-28',str(today))
            dsa('P_age', values['age'])
            dsa('P_gen', values['gender'])
            dsa('8976543210', values['contact'])
            dsa('100', values['temp'])
            dsa('110', values['bp'])
            dsa('98', values['spo2'])
            dsa('ABCDEF', str(y_find[0]))
            
            
        #doc.save('output.pdf')
        doc.save('output.docx')
        docx2pdf.convert('output.docx','output.pdf')
        #doc.remove('output.docx')
        webbrowser.open_new_tab('output.pdf')


# Close the window when the loop exits
window.close()
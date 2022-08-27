
import PySimpleGUI as sg
sg.theme('BluePurple')
layout = [[sg.Text('Nhập đường dẫn file input'),
           sg.Text(size=(15,1), key='output1')],
            [sg.Input(key='input1')],
           [sg.Text('Nhập đường dẫn file output :'),
           sg.Text(size=(15,1), key='output2')],
            [sg.Input(key='input2')],
          [sg.Button('Run'), sg.Button('Exit')]]
  
window = sg.Window('Introduction', layout)
  
while True:
    event, values = window.read()
    print(event, values)
      
    if event in  (None, 'Exit'):
        break
      
    if event == 'Run':
        # Update the "output" text element
        # to be the value of "input" element
        window['output1'].update(values['input1'])
        window['output2'].update(values['input2'])
  
window.close()
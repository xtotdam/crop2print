import PySimpleGUI as sg
import PyPDF2

import dataclasses
from pathlib import Path
import subprocess
import webbrowser

@dataclasses.dataclass
class S:
    ''' settings '''
    init_margins = dict(t='10', b='10', l='15', r='15') # PhysRev
    theme = 'Reddit'
    sumatra_path = Path(r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe")

sg.theme(S.theme)

layout = [
    [
        sg.Text('PDF: '), sg.In(enable_events=True, key='filename'),
        sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),("All Files", "*.*")), initial_folder=Path().cwd())
    ],
    [sg.Text('', size=(40,2), key='pdfinfo')],

    [
        sg.Frame('Crop margins (mm)',
            [
                [sg.Text('', size=(5, 1)), sg.InputText(S.init_margins['t'], size=(5, 1), key='crop-t', tooltip='Top margin')],
                [
                    sg.InputText(S.init_margins['l'], size=(5, 1), key='crop-l'),
                    sg.Text('', size=(5, 1)),
                    sg.InputText(S.init_margins['r'], size=(5, 1), key='crop-r')
                ],
                [sg.Text('', size=(5, 1)), sg.InputText(S.init_margins['b'], size=(5, 1), key='crop-b')]
            ]),
        sg.Column([
            [sg.Button('Crop', size=(7,1))],
            [sg.Button('Open original in Sumatra', key='sumatraorig', size=(25,1))],
            [sg.Button('Open cropped in Sumatra', key='sumatra', size=(25,1))],
            ]),
    ],

    [sg.Text('', size=(49,1)), sg.Text("by Xtotdam", text_color='blue', enable_events=True, key='githublink')]
]

window = sg.Window('Crop 2 Print', layout, margins=(0,0), grab_anywhere=True, keep_on_top=True)
window.Finalize()

while True:  # Event Loop
    event, values = window.read()       # can also be written as event, values = window()
    # print(event, values)
    if event is None or event == 'Exit':
        break

    elif event == 'githublink':
        webbrowser.open('https://github.com/xtotdam/crop2print')

    elif event == 'filename':
        pdf_src = Path(values['filename'])

        if pdf_src.exists() and len(values['filename']) > 0:
            window['filename'].update(background_color='#aaffaa')

            pdfReader = PyPDF2.PdfFileReader(open(pdf_src, 'rb'))
            pages_have_same_size = True
            mb0 = pdfReader.getPage(0).mediaBox
            for n in range(1, pdfReader.numPages):
                mb1 = pdfReader.getPage(n).mediaBox
                if sum(mb0[i] == mb1[i] for i in range(4)) != 4:
                    pages_have_same_size = False
                    break

            papersize = list(pdfReader.getPage(0).mediaBox.upperRight)
            papersize[0] = float(papersize[0]) / 72 * 25.4
            papersize[1] = float(papersize[1]) / 72 * 25.4

            pdfinfo = f'{pdfReader.numPages} pages. First page is {papersize[0]:.1f} x {papersize[1]:.1f} mm\nPages have same size: {pages_have_same_size}'

            window['pdfinfo'].update(pdfinfo)
        else:
            window['filename'].update(background_color='#ffaaaa')
            window['pdfinfo'].update('')

    elif event == 'Crop':
        if pdf_src.exists() and len(values['filename']) > 0:
            pdf_src = Path(values['filename'])
            pdf_out = pdf_src.parent / (pdf_src.stem + '-c2p.pdf')

            pdfReader = PyPDF2.PdfFileReader(open(pdf_src, 'rb'))
            pdfWriter = PyPDF2.PdfFileWriter()

            for n in range(pdfReader.numPages):
                pdfWriter.addPage(pdfReader.getPage(n))
                page = pdfWriter.getPage(n)

                page.mediaBox.lowerLeft  = (
                    (float(page.mediaBox.lowerLeft[0]) / 72 * 25.4 + float(values['crop-l'])) / 25.4 * 72,
                    (float(page.mediaBox.lowerLeft[1]) / 72 * 25.4 + float(values['crop-b'])) / 25.4 * 72
                )

                page.mediaBox.upperRight = (
                    (float(page.mediaBox.upperRight[0]) / 72 * 25.4 - float(values['crop-r'])) / 25.4 * 72,
                    (float(page.mediaBox.upperRight[1]) / 72 * 25.4 - float(values['crop-t'])) / 25.4 * 72
                )

            with open(pdf_out, 'wb') as out:
                pdfWriter.write(out)

    elif event == 'sumatra':
        pdf_src = Path(values['filename'])
        pdf_out = pdf_src.parent / (pdf_src.stem + '-c2p.pdf')
        if pdf_out.exists():
            cmd = [S.sumatra_path, pdf_out]
            p = subprocess.Popen(cmd, shell=False, stdin=None, stdout=None, stderr=None, close_fds=True, creationflags=subprocess.DETACHED_PROCESS)

    elif event == 'sumatraorig':
        pdf_src = Path(values['filename'])
        if pdf_src.exists() and len(values['filename']) > 0:
            cmd = [S.sumatra_path, pdf_src]
            p = subprocess.Popen(cmd, shell=False, stdin=None, stdout=None, stderr=None, close_fds=True, creationflags=subprocess.DETACHED_PROCESS)

window.close()

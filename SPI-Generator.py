import re

import openpyxl
from sys import stderr
import os.path
import pathlib
import types
import streamlit as st
from tempfile import NamedTemporaryFile


rfc_path = pathlib.Path(r"C:\Users\nick\Downloads\Desktop")
print(rfc_path) 



def get_reimbursable_pay_items():
    reimbursable_pay_items = ["1-4.","1-6.","2-3.","4-1.b.","14-3.","15-3.","15-3.a.","15-3.b.","15-7.",
                          "15-8.","15-9.","15-10.","22-1.","22-2.","22-3.","22-4.","22-5.","22-6.",
                          "22-7.","22-8.","22-9.","22-10.","22-11.","22-12.","22-13.","22-14.","22-15.",
                          "22-16.","22-17.","22-18.","22-19.","22-20.","22-21.","22-22.","22-23.","22-24.",
                          "22-25.","22-26.","22-27.","22-28.","22-29.","22-30.",]
    return reimbursable_pay_items

def my_input(prompt=None):
    '''this function replaces input() and prints prompt to stderr'''
    if prompt:
        stderr.write(str(prompt))
    return input()
    
def load_rfc(rfc_worksheet):
    '''this takes a worksheet from RFC file and creates a list of tuples, 
    each tuple includes tsk, pi, uom, pi_price, new_val, total_price'''
    lol = []
    for row in rfc_worksheet.iter_rows(min_row=22, max_row=111, min_col=0,max_col=9, values_only=True):
        if row[0] == None: continue
        tsk, pi, desc,_,_, uom, pi_price, new_val, total_price = row
        lol.append(('TASK '+str(tsk),pi, uom, new_val, total_price))
    lol.sort()
    return lol

def get_spi(spi_wb):
    """This function creates an instance of the SPI, this keeps the included vba and the links intact"""
    print(spi_wb)
    return openpyxl.load_workbook(spi_wb, keep_vba=True, keep_links=True)

def get_rfc(rfc_wb):
    """This function creates an instance of the RFC Workbook object, it also ensures that 
    its in read only mode and data only mode for speed and data integrity"""
    print(rfc_wb)
    return openpyxl.load_workbook(rfc_wb, read_only=True, data_only=True, keep_links=True)

_LABEL_RE = re.compile(r'^\d{1,2}([-.][0-9A-Za-z]+)*\.$')
_TASK_RE = re.compile(r'TASK\s*(\d+)', re.IGNORECASE)
_LAYOUT_CACHE = {}


def spi_layout(ws):
    '''Locate the pay item table in an SOW Units worksheet from the
    workbook itself. The header row carries a 'PAY ITEM' cell and the
    'TASK n QUANT.' column headers; pay item labels run down the PAY ITEM
    column below the header, ending at the totals/version block. Returns
    (labels in workbook order, {label: row}, {task number: column}).
    Raises ValueError when the table cannot be found -- positions are
    never guessed, because a workbook from a different SPI version has
    different rows (Version 10.0, for example, has no section 22, so
    section 23 sits 31 rows above where a fixed list would look).'''
    key = id(ws)
    if key in _LAYOUT_CACHE:
        return _LAYOUT_CACHE[key]
    header = None
    for row in ws.iter_rows(min_row=1, max_row=40, max_col=24):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower() == 'pay item':
                header = (cell.row, cell.column)
                break
        if header:
            break
    if not header:
        raise ValueError('no pay item table found in the SPI workbook: '
                         'no "PAY ITEM" header cell in the SOW Units sheet')
    hr, lc = header
    task_cols = {}
    for row in ws.iter_rows(min_row=max(1, hr - 1), max_row=hr + 1, max_col=24):
        for cell in row:
            if isinstance(cell.value, str):
                m = _TASK_RE.search(cell.value)
                if m:
                    task_cols[int(m.group(1))] = cell.column
    if not task_cols:
        raise ValueError('no task quantity columns found in the SPI workbook')
    labels, rows, misses, started = [], {}, 0, False
    label_rows = ws.iter_rows(min_row=hr + 1, max_row=ws.max_row,
                              min_col=lc, max_col=lc, values_only=True)
    for offset, (v,) in enumerate(label_rows):
        s = str(v).strip() if v is not None else ''
        if _LABEL_RE.match(s):
            started, misses = True, 0
            if s not in rows:
                rows[s] = hr + 1 + offset
                labels.append(s)
        elif started:
            misses += 1
            if misses >= 10:
                break
    if len(labels) < 10:
        raise ValueError('only {} pay items found in the SPI workbook'.format(len(labels)))
    layout = (labels, rows, task_cols)
    _LAYOUT_CACHE[key] = layout
    return layout


def get_old_spi_cell(ws, pi, task):
    '''The cell for a pay item and task, located from the workbook's own
    pay item table. Raises ValueError naming the item when the workbook
    has no row with that label.'''
    labels, rows, task_cols = spi_layout(ws)
    if pi not in rows:
        raise ValueError('pay item "{}" is not in the SPI workbook: its pay '
                         'item table has no row with that label'.format(pi))
    tnum = int(str(task).split()[-1])
    if tnum not in task_cols:
        raise ValueError('task "{}" has no quantity column in the SPI workbook'.format(task))
    cell = ws.cell(row=rows[pi], column=task_cols[tnum])
    cell.value = 0 if isinstance(cell.value, types.NoneType) else cell.value
    return cell

@st.fragment
def my_download_button(file):
    st.download_button(label='Download Generated SPI', 
                data=file, file_name="newlygeneratedspi.xlsm", mime="application/vnd.ms-excel", 
                key=None, help=None, on_click=None, 
                args=None, kwargs=None, type="primary", 
                disabled=False, use_container_width=False)
    
def _norm_qty(v):
    # Quantities in this program are decimal values -- dollar amounts to the
    # cent, and FDEP-issued quantities carry at most 6 decimal places. A raw
    # binary-float sum can land just off the intended decimal (1000.0 +
    # -999.99 = 0.009999999999990905 instead of 0.01). Round each computed
    # quantity to 6 decimals so the workbook holds the intended value.
    r = round(v, 6)
    return int(r) if float(r).is_integer() else r


def generate_report(rfc_items, spi, SOW_Units):
    report=[]
    success = None
    try:
        for rfc_item in rfc_items:
            task,pi,units = rfc_item[0],rfc_item[1],rfc_item[3]
            oldval = get_old_spi_cell(SOW_Units,pi, task).value
            get_old_spi_cell(SOW_Units, pi, task).value += units
            get_old_spi_cell(SOW_Units, pi, task).value = _norm_qty(get_old_spi_cell(SOW_Units, pi, task).value)

            report.append('{} unit(s) added to {} in {}, {} The new value is {}'.format(
                units, 
                oldval, 
                task,
                pi, 
                get_old_spi_cell(SOW_Units, pi, task).value))
            print('{} unit(s) added to {} in {}, {} The new value is {}'.format(
                units, 
                oldval, 
                task,
                pi, 
                get_old_spi_cell(SOW_Units, pi, task).value))
            SOW_Units.row_dimensions[spi_layout(SOW_Units)[1][pi]].hidden = False
        success = True
        spi.save('./newgeneratedspi.xlsm')
    except ValueError:
        # Layout errors (pay item table not found, RFC item missing from
        # the workbook) name the problem; do not mask them as a generic
        # failure or write a file.
        raise
    except:
        report.append('Something went wrong...')
        success = False
    else:
        report.append('### Transfer Complete')
    finally:
        pass

        
    return report, success

def main():
    st.write('# Generate SPI from Change Order')
    st.write('''### This tool transfers pay items to the new SPI''')
    
    col1, col2 = st.columns(2)

    with col1:
        st.write('### Step 1.')
        rfc = st.file_uploader("Upload your Excel RFC here",type='xlsm')
    
    with col2:
        st.write('### Step 2.')
        old_spi = st.file_uploader("Upload Old SPI",type='xlsm')
    if rfc!=None and old_spi!=None:
        
        with st.spinner('Wait while processing...'):
            wb_rfc = get_rfc(rfc)
            rfc_worksheet = wb_rfc['RFC']
            spi = get_spi(old_spi)
            SOW_Units = spi['SOW Units']

    try:
        report, success = generate_report(load_rfc(rfc_worksheet), spi, SOW_Units)
        report = '  \n'.join(report)
        if success:
            st.image('https://www.nicepng.com/png/full/362-3624869_success-image-png.png',width=100,)
            st.write(report)
     
        else:
            st.image('https://purepng.com/public/uploads/large/purepng.com-sign-stoptraffic-signsign-stopnotify-driversstop-signs-1701527614263a3116.png',width=100)

    except:
        st.write('Upload your files to process')
    else:
        st.write('### Step 3.')   
        file = open("./newgeneratedspi.xlsm", "rb")
        my_download_button(file)    

    finally: 
        st.caption('''Like this tool? Check out the [Change Order Checker](https://change-order-checker.streamlit.app/)''')
        st.caption('''This tool is authored by Nicholas Moran PG at [Moran.Rocks](https://Moran.Rocks)''')
        st.caption('''Please Share on [Linkedin](https://www.linkedin.com/shareArticle?mini=true&url=https%3A//spigenerator.streamlit.app/) or send an [email](mailto:?body=Check%20out%20the%20new%20SPI%20generator%20https%3A//spigenerator.streamlit.app/)''')
   
if __name__=="__main__":
    main()

import React, { Component } from 'react';
import XLSX from 'xlsx';
import { make_cols } from './MakeColumns';

const SheetJSFT = [
  'xlsx',
  'xlsb',
  'xlsm',
  'xls',
  'xml',
  'csv',
  'txt',
  'ods',
  'fods',
  'uos',
  'sylk',
  'dif',
  'dbf',
  'prn',
  'qpw',
  '123',
  'wb*',
  'wq*',
  'html',
  'htm',
]
  .map(function (x) {
    return '.' + x;
  })
  .join(',');

class ExcelReader extends Component {
  constructor(props) {
    super(props);
    this.state = {
      file: {},
      data: [],
      cols: [],
      result: [],
    };
    this.handleFile = this.handleFile.bind(this);
    this.handleChange = this.handleChange.bind(this);
  }

  handleChange(e) {
    const files = e.target.files;
    if (files && files[0]) this.setState({ file: files[0] });
  }

  handleFile() {
    /* Boilerplate to set up FileReader */
    const reader = new FileReader();
    const rABS = !!reader.readAsBinaryString;

    reader.onload = (e) => {
      /* Parse data */
      const bstr = e.target.result;
      const wb = XLSX.read(bstr, {
        type: rABS ? 'binary' : 'array',
        bookVBA: true,
      });
      /* Get first worksheet */
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      /* Convert array of arrays */
      const data = XLSX.utils.sheet_to_json(ws);
      /* Update state */
      this.setState({ data: data, cols: make_cols(ws['!ref']) }, () => {
        this.setState({ result: JSON.stringify(this.state.data, null, 2) });
        // console.log(JSON.stringify(this.state.data, null, 2));
      });
    };

    if (rABS) {
      reader.readAsBinaryString(this.state.file);
    } else {
      reader.readAsArrayBuffer(this.state.file);
    }
  }

  render() {
    return (
      <div>
        <label htmlFor='file'>Upload an excel sheet to convert to JSON</label>
        <br />
        <br />
        <input
          type='file'
          className='form-control'
          id='file'
          accept={SheetJSFT}
          onChange={this.handleChange}
        />
        <br />
        <br />
        <input type='submit' value='Show JSON' onClick={this.handleFile} />
        <hr />
        <h3>Result:</h3>
        <pre
          style={{
            width: 'auto',
            border: '2px solid black',
            borderRadius: '5px',
            margin: 'auto 25px',
            padding: '10px',
          }}
        >
          {this.state.result == '' ? 'No Output' : this.state.result}
        </pre>
      </div>
    );
  }
}

export default ExcelReader;

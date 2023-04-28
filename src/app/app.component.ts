import { Component, OnInit } from '@angular/core';


import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'DataComparison';

 
  
  workingDay: any = 0;

  _data1: any = [];
  _data2: any = [];
  _data3: any = [];

  _cgData: any = [];
  _subaruData: any = [];
  _leaveData: any = [];


  selectedOption: String = " ";
  _uniqueMgn: any = [];

  _mergeData: any = [];

  _exportData: any[] = [];









  // for sheet1 excel data convert into JSON
  cgFileUpload1(event: any) {
    const _selectedFile = event.target.files[0];

    const _fileReader = new FileReader();
    _fileReader.readAsBinaryString(_selectedFile);

    _fileReader.onload = (event) => {
      let _binaryData = event.target?.result;

      let _workbook = XLSX.read(_binaryData, { type: 'binary' });

      _workbook.SheetNames.forEach((sheet) => {
        const _data = XLSX.utils.sheet_to_json(_workbook.Sheets[sheet]);

        // let _dataW = JSON.stringify(_data, undefined, 4);

        this._data1 = _data;
        // console.log("data 1",this._data1);

        this.formatDataforExcel1();
      });
    };
  }



  // for sheet2 excel data convert into JSON
  sbFileUpload2(event: any) {
    const _selectedFile = event.target.files[0];

    const _fileReader = new FileReader();
    _fileReader.readAsBinaryString(_selectedFile);

    _fileReader.onload = (event) => {
      let _binaryData = event.target?.result;

      let _workbook = XLSX.read(_binaryData, { type: 'binary' });

      _workbook.SheetNames.forEach((sheet) => {
        const _data = XLSX.utils.sheet_to_json(_workbook.Sheets[sheet]);

        this._data2 = _data;
        // console.log("data2",this._data2);
        this.formatDataforExcel2();
      });
    };
  }



  // for sheet3 excel data convert into JSON
  leavefileUpload3(event: any) {
    const _selectedFile = event.target.files[0];

    const _fileReader = new FileReader();
    _fileReader.readAsBinaryString(_selectedFile);

    _fileReader.onload = (event) => {
      let _binaryData = event.target?.result;

      let _workbook = XLSX.read(_binaryData, { type: 'binary' });

      _workbook.SheetNames.forEach((sheet) => {
        const _data = XLSX.utils.sheet_to_json(_workbook.Sheets[sheet]);

        this._data3 = _data;

        // console.log("leave Data",this._data3);
        this.formatDataforExcel3();

      });
    };
  }




  // after convert excel data file 1 into JSON and do operation according instruction on JSON data of array    
  formatDataforExcel1() {
    const _max1 = this._data1;

    const _employee: any = [];

    _max1.forEach((item: any) => {
      let obj: any = {};
      if (item["Task ID"] == 'BT') {
        obj['Employee Name'] = item['Employee Name'];
        obj['Billable'] = item.Billable;
        obj['Hrs'] = item.Hrs;
        _employee.push(obj);
      }
    });

    const _realName = _employee.map((item: any) => item['Employee Name']);

    const _filterName = this.removeDups(_realName);

    const cgTime: any = [];

    _filterName.forEach((item: any) => {
      let sum = 0,
        obj1: any = {};
      for (let i = 0; i < _employee.length; i++) {
        if (item == _employee[i]['Employee Name']) {
          sum += _employee[i]['Hrs'];
        }
      }
      obj1['Employee Name'] = item;
      obj1['CG_Hours'] = sum;
      cgTime.push(obj1);
    });

    this._cgData = cgTime;
    // console.log("CG data",this._cgData)
  }


  // after convert excel data file 2 into JSON and do operation according instruction on JSON data of array    
  formatDataforExcel2() {
    const _max2 = this._data2;
    const _employeeName: any = [];

    _max2.forEach((item: any) => {
      let obj: any = {};
      if (item.Manager == this.selectedOption) {
        obj['User'] = item.User;
        obj['Total Hours'] = item['Total Hours'];
        _employeeName.push(obj);
      }
    });

    const _realName = _employeeName.map((item: any) => item.User);
    const _filterName = this.removeDups(_realName);

    const _realNameMng = this._data2.map((item: any) => item.Manager);
    const _filterNameMng = this.removeDups(_realNameMng);
    this._uniqueMgn = _filterNameMng;

    const subaruTime: any = [];

    _filterName.forEach((item: any) => {
      let sum = 0,
        obj1: any = {};
      for (let i = 0; i < _employeeName.length; i++) {
        if (item == _employeeName[i].User) {
          sum += _employeeName[i]['Total Hours'];
        }
      }
      obj1['User'] = item;
      obj1['Subaru_Hours'] = sum;
      subaruTime.push(obj1);
    });

    this._subaruData = subaruTime;
    // console.log("subaru data",this._subaruData);

  }



  // after convert excel data file 3 into JSON and do operation according instruction on JSON data of array    
  formatDataforExcel3() {
    const _max3 = this._data3;

    const _employee: any = [];

    _max3.forEach((item: any) => {
      let obj: any = {};
      if (item['Employee Name'] != undefined) {

        obj['Employee Name'] = item['Employee Name'];
        obj['Leave Status'] = item['Leave Status'];
        _employee.push(obj);

      }
    });

    // for remove duplicate name from array of object
    const _realName = [...new Set(_employee.map((item: any) => item['Employee Name']))];

    const leaveTime: any = [];

    _realName.forEach((item: any) => {
      let sum = 0,
        obj1: any = {};
      for (let i = 0; i < _employee.length; i++) {
        if (item == _employee[i]['Employee Name']) {
          sum += 8;
        }
      }
      obj1['Employee Name'] = item;
      obj1['Leave_Hours'] = sum;
      leaveTime.push(obj1);
    });


    this._leaveData = leaveTime;
    // console.log(this._leaveData, 'leaveTime data');


  }




  // for filter out data accourding option in subarau sheet  
  change(event: any) {
    this.selectedOption = event;
    this.formatDataforExcel2();

  }



  // for form sumbit
  onSubmit() {
    this.mergeData1()

    // console.log("data is submitted.........");

  }





  // merge array of CG data and Subaru data after final optimization and this filna output
  mergeData1() {
    this._exportData = [];

    if (this._cgData.length == this._subaruData.length) {
      this._cgData.forEach((item: any, index: number) => {
        let obj: any = {};

        obj['Sr No'] = index + 1;
        obj['Employee name'] = item['Employee Name'];
        obj['Working Hours'] = this.workingDay * 8;
        obj['CG Hours'] = item.CG_Hours;
        obj['Leave Hours'] = this.getLeaveData(item['Employee Name']) != undefined ? this.getLeaveData(item['Employee Name']) : 0;

        obj['Subaru Hours'] = this.getSubaruData(item['Employee Name']) != undefined ? this.getSubaruData(item['Employee Name']) : 0;

        obj['Total Hours'] = this.getTotalHours(item) != undefined ? this.getTotalHours(item) : 0;

        obj['Data Miss Match reason'] = this.getMatchData(item);
        this._exportData.push(obj);
      });

      // console.log(this._exportData, 'exportData');
    } else {
      window.alert("please select right manager...")
    }

  }




  getMatchData(data: any) {

    let match = "";
    let notMatch = "";
    let dataMatch = "";

    const formatData1 = data["Employee Name"].replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
    const dFirstName = formatData1.split(" ")[0];
    const dLastName = formatData1.split(" ")[1];

    this._subaruData.find((item: any) => {

      const formatData2 = item.User.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
      const sbFirstName = formatData2.split(" ")[0];
      const sbLastName = formatData2.split(" ")[1];

      if (dFirstName == sbFirstName || dFirstName == sbLastName || dLastName == sbFirstName || dLastName == sbLastName) {

        dataMatch = data.CG_Hours == item.Subaru_Hours ? "accurate" : "CG and Subaru are not matched";
      }
    }
    );

    return dataMatch;

  }




  getTotalHours(data: any) {
    let totalHour = 0;

    const formatData1 = data['Employee Name'].replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
    const dFirstName = formatData1.split(" ")[0];
    const dLastName = formatData1.split(" ")[1];


    this._leaveData.find((item: any) => {

      const formatData2 = item['Employee Name'].replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
      const lFirstName = formatData2.split(" ")[0];
      const lLastName = formatData2.split(" ")[1];

      if (dFirstName == lFirstName || dFirstName == lLastName || dLastName == lFirstName || dLastName == lLastName) {
        totalHour = data.CG_Hours + item.Leave_Hours;
      }
    }
    );

    if (totalHour == 0) {
      return data.CG_Hours;
    }
    else {

      return totalHour;
    }


  }




  getLeaveData(data: string) {
    let leaveHour = 0;

    const formatData1 = data.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
    const dFirstName = formatData1.split(" ")[0];
    const dLastName = formatData1.split(" ")[1];


    this._leaveData.find((item: any) => {

      const formatData2 = item['Employee Name'].replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
      const lFirstName = formatData2.split(" ")[0];
      const lLastName = formatData2.split(" ")[1];

      if (dFirstName == lFirstName || dFirstName == lLastName || dLastName == lFirstName || dLastName == lLastName) {
        leaveHour = item.Leave_Hours;
      }
    }
    );

    return leaveHour;
  }


  getSubaruData(data: string) {
    let subaruHour = 0;

    const formatData1 = data.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
    const dFirstName = formatData1.split(" ")[0];
    const dLastName = formatData1.split(" ")[1];

    this._subaruData.find((item: any) => {

      const formatData2 = item.User.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').toUpperCase().trim();
      const sbFirstName = formatData2.split(" ")[0];
      const sbLastName = formatData2.split(" ")[1];

      if (dFirstName == sbFirstName || dFirstName == sbLastName || dLastName == sbFirstName || dLastName == sbLastName) {
        subaruHour = item.Subaru_Hours;
      }
    }
    );

    return subaruHour;
  }








  // this function for remove duplicate name from array of object
  removeDups(names: any) {

    let unique: any = {};
    names.forEach(function (i: any) {
      if (!unique[i]) {
        unique[i] = true;
      }
    });
    return Object.keys(unique);
  }





  // this function for ouput data array of object convert into excel form
  exportFileData(): void {
    this.exportArrayToExcel(this._exportData, 'new_File')
  };
  exportArrayToExcel(arr: any[], name: string) {
    let { sheetName, fileName } = this.getFileName(name);
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.json_to_sheet(arr);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${fileName}.xlsx`);
  }



  // this function for current data with file name at download time save file 
  getFileName = (name: string) => {
    let timeSpan = new Date().toISOString();
    let sheetName = name || "ExportResult";
    let fileName = `${sheetName}-${timeSpan}`;
    return {
      sheetName,
      fileName
    };
  };



}

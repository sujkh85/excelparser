import XLSX from 'xlsx';
import fileSaver from 'file-saver';
import isJson from 'is-json'

class ExcelUtil {
  static download2(strData, strFileName, strMimeType='text/plain') {
    let a = document.createElement("a")
    a.setAttribute("download", strFileName);
    a.setAttribute('href', `data:${strMimeType};charset=utf-8, ${strData}`)
    a.click()
  }
    static download(strData, strFileName, strMimeType) {
        var D = document,
            A = arguments,
            a = D.createElement("a"),
            d = A[0],
            n = A[1],
            t = A[2] || "text/plain";

        //build download link:
        a.href = "data:" + strMimeType + "charset=utf-8," + escape(strData);


        // if (window.MSBlobBuilder) { // IE10
        //     var bb = new MSBlobBuilder();
        //     bb.append(strData);
        //     return navigator.msSaveBlob(bb, strFileName);
        // } /* end if(window.MSBlobBuilder) */



        if ('download' in a) { //FF20, CH19
            a.setAttribute("download", n);
            a.innerHTML = "downloading...";
            D.body.appendChild(a);
            setTimeout(function() {
                var e = D.createEvent("MouseEvents");
                e.initMouseEvent("click", true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
                a.dispatchEvent(e);
                D.body.removeChild(a);
            }, 66);
            return true;
        }; /* end if('download' in a) */



        //do iframe dataURL download: (older W3)
        var f = D.createElement("iframe");
        D.body.appendChild(f);
        f.src = "data:" + (A[2] ? A[2] : "application/octet-stream") + (window.btoa ? ";base64" : "") + "," + (window.btoa ? window.btoa : escape)(strData);
        setTimeout(function() {
            D.body.removeChild(f);
        }, 333);
        return true;
    }
  static _datenum(v, date1904) {
    if(date1904) v+=1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  static _sheet_from_array_of_arrays(data, opts){
    var ws = {};
    var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
    for(var R = 0; R !== data.length; ++R) {
      for(var C = 0; C !== data[R].length; ++C) {
        if(range.s.r > R) range.s.r = R;
        if(range.s.c > C) range.s.c = C;
        if(range.e.r < R) range.e.r = R;
        if(range.e.c < C) range.e.c = C;
        var cell = {v: data[R][C] };
        if(cell.v == null) continue;
        var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
        
        if(typeof cell.v === 'number') cell.t = 'n';
        else if(typeof cell.v === 'boolean') cell.t = 'b';
        else if(cell.v instanceof Date) {
          cell.t = 'n'; cell.z = XLSX.SSF._table[14];
          cell.v = this._datenum(cell.v);
        }
        else cell.t = 's';
        
        ws[cell_ref] = cell;
      }
    }
    if(range.s.c < 10000000){
      ws['!ref'] = XLSX.utils.encode_range(range);
    }
    return ws;
  }

  static _s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!==s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  static _createWorkbook() {
    return {SheetNames:[], Sheets:{}};
  }

  /*
    example
    var fileName = '파일명.xlsx'
    var data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]
    var title = "SheetJS";
    write(title, data)
  */

  static getUnlockFormat(fileName, title, callback){
    let data = [['txHash'],['0x068c55b2b43421b4e66188de9fce20bea49b26a7935e4c6169343e4ce71943af'],['0x29d865f8fe3594e290a6009f47498ab57706cfce022a98ffc10269d453409fb2']]
    this.write(fileName, title, data, callback)
  }

  static getWhiteListFormat(fileName, title, callback){
    let data = [
      ['from','min','max'],
      ['0xc98636822709323660e429e015D4E739BD25dab5',0,10],
      ['0x9717Df991e66f18411056Ae8186f1aC2Fa5c6299',3,20],
      ['0x627306090abaB3A6e1400e9345bC60c78a8BEf57',1,10],
      ['0xA70Dff406A82344c43F8416f5530a6Cd0FD8F83D',0,11]
    ]
    this.write(fileName, title, data, callback)
  }

  static write(fileName='test.xlsx', title='sheet', data=[[1,2,3]], callback){
    var wb = this._createWorkbook(), ws = this._sheet_from_array_of_arrays(data);
 
    /* add worksheet to workbook */
    wb.SheetNames.push(title);
    wb.Sheets[title] = ws;
    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});

    fileSaver.saveAs(new Blob([this._s2ab(wbout)],{type:"application/octet-stream"}), fileName)
    
    if(callback){
      callback()
    }
  }

  static xlsxReader(files, xlangth){
    return new Promise((resolve, reject)=>{
      if(files.length === 0){
        reject('length 0')
      }
      //let fileNameObject = this.getFileNameObject(files[0]);
      let file = files[0];
      
      let fileReader = new FileReader();
      fileReader.onload =(e)=>{
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, {type:'array'})
        let result
        result = ExcelUtil.xlsxPaser(workbook, xlangth)
        
        resolve(result)  
      }
      fileReader.readAsArrayBuffer(file)  
      
    })
  }
  static _diff(ref,xlangth=5){
    let split = ref.split(':')
    let xStart = split[0].substring(0,1).charCodeAt()
    let xEnd = split[0].substring(1,split[0].length)*1
    let yStart = split[1].substring(0,1).charCodeAt()
    let yEnd = split[1].substring(1,split[1].length)*1
    // let xDiff = yStart - xStart
    let xDiff = xlangth
    let yDiff = yEnd - xEnd
    let xKeyList= [] 
    for(let i=0;i<=xDiff;i++){
        let xStartKey = String.fromCharCode(xStart+i)
        xKeyList.push(xStartKey)
    }
    return {xDiff, yDiff, xKeyList}
  }
  static replaceAll(str, searchStr, replaceStr) {
    return str.split(searchStr).join(replaceStr);
  }
  static makeFileObj(Sheets,diff){
    let max = diff.yDiff+1
    let fileObj = {} 
    diff.xKeyList.forEach((key,index)=>{
        let fileName =''
        for(let i=1;i<=max;i++){
          //파일 이름 세팅
            if(i === 1 && index !== 0){
                if(Sheets[key+i] && Sheets[key+i].w){
                    fileName=Sheets[key+i].w
                    fileObj[fileName]={}
                }
            }else{
              //파일명을 기준으로 데이터 생성
                if(index !==0){
                    let targetKey = Sheets['A'+i].w
                    if(Sheets[key+i]){
                        let target = Sheets[key+i].w
                        //배열인지 확인
                        if(target.indexOf('[') !== -1 && target.indexOf(']') !== -1){
                            let comma= ExcelUtil.replaceAll(target,'"','') 
                            let a= comma.replace('[','')
                            let b= a.replace(']','')
                            let split = b.split(',')
                            let result = []
                            split.forEach((item)=>{
                              let replaced = item
                                result.push(replaced)
                            })
                            //중복된 key값인지 확인
                            if(fileObj[fileName][targetKey]){
                              alert('중복된키값 : '+targetKey)
                            }
                            fileObj[fileName][targetKey]=result
                            
                        }
                        //실제 데이터 확인
                        else{
                          let replaced = target
                          if(fileObj[fileName][targetKey]){
                            alert('중복된키값 : '+targetKey)
                          }
                          fileObj[fileName][targetKey]=replaced  
                        }
                    }else{
                        fileObj[fileName][targetKey]=''
                    }
                }
            }
        }
    })
    return fileObj
  }
  static xlsxPaser(workbook,xlangth){
    let SheetNames = workbook.SheetNames[0]
    let Sheets = workbook.Sheets[SheetNames]
    let diff = ExcelUtil._diff(Sheets['!ref'],xlangth)
    let result = {}
    let fileObj = ExcelUtil.makeFileObj(Sheets,diff)
    return fileObj
  }
  static ObjectToArr(obj) {
    if (!obj) {
      return [];
    }
    let arr = [];
    Object.keys(obj).forEach(key => {
      arr.push(obj[key]);
    });
    return arr;
  }
}

export default ExcelUtil
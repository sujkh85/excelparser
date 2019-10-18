import React, { Component } from 'react';
import './App.css';
import ExcelUtil from './ExcelUtil'

class App extends Component {
  constructor(props){
    super(props)
    this.state = {
      
    }
  }
  onChangeFile=(e)=>{
    let languageCount = 5
    ExcelUtil.xlsxReader(e.target.files, languageCount).then((parse)=>{
      console.log('parse', parse)
      this.setState({
        language:parse
      })
      // let strMimeType = 'application/json'
      // Object.keys(parse).forEach((fileName,index)=>{
      //   let json = JSON.stringify(parse[fileName])
      //   setTimeout(()=>{
      //     ExcelUtil.download(json,fileName,strMimeType)
      //   }
      //   ,index*1000)
      // })
    })
  }
  onClickLanguage=(e)=>{
    let fileName = e.currentTarget.getAttribute('data')
    // let strMimeType = 'application/json'
    if(fileName === 'All'){
      let {language} = this.state
      Object.keys(language).forEach((key,index)=>{
        let json = JSON.stringify(language[key])
        setTimeout(()=>{
          ExcelUtil.download2(json,key)
        }
        ,index*1200)
      })
    }else{
      console.log('fileName',fileName)
      let {language} = this.state
      console.log('language', language)
      if(language){
        let target = language[fileName+'.json']
        console.log('target', target)
        // let strMimeType = 'application/json'
        let jsonString = JSON.stringify(target)
        
        // ExcelUtil.replaceAll(jsonString,'%u2194',"↔")
        // ExcelUtil.replaceAll(jsonString,'%u203B',"※")
        // ExcelUtil.replaceAll(jsonString,'%u2018',"‘")
        // ExcelUtil.replaceAll(jsonString,'%u2019',"’") 
        console.log('jsonString', jsonString)
        ExcelUtil.download2(jsonString,fileName+'.json')
      }
    }
    
  }
  LanguageList = ['English', 'Korean','Chinese','Vietnamese','Russian']
  render() {
    const {language}=this.state
    return (
      <div className="App">
        <input type="file" onChange={this.onChangeFile}></input>
        {language && <div style={{display:'flex', marginTop:20}}>
          {this.LanguageList.map((language)=>{
            return (
              <div className="flex-button" onClick={this.onClickLanguage} data={language} key={language}>{language}</div>    
            )
          })}
          <div className="flex-button" onClick={this.onClickLanguage} data="All">All</div>
        </div>}

      </div>
    );
  }
}

export default App;

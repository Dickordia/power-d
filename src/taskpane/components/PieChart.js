/* global Button, console, Header, HeroList, HeroListItem, Office, Progress */

import * as React from "react";

import * as d3 from "d3";
import ReactDOM from "react-dom";

import PieClass from "./PieClass";
import { Button, ButtonType } from "office-ui-fabric-react";
const kLocation = "https://dickordia.github.io/power-d"

import Lottie from 'lottie-react-web'
import animation from './anim3.json'

import ReactDataSheet from 'react-datasheet';
import './tableStyle.css';

import { CirclePicker } from 'react-color';
import Popup from "reactjs-popup";
import 'react-datasheet/lib/react-datasheet.css';
import './PieChart.css'

var dialog = null;
var chart = null;
var aColors = d3.schemeCategory10;
var data = [{date: 0, value: 36, color: aColors[0]},
            {date: 1, value: 17, color: aColors[1]},
            {date: 2, value: 54, color: aColors[2]},
            {date: 3, value: 26, color: aColors[3]},
            {date: 4, value: 76, color: aColors[4]}]

function processMessage(arg) {
  chart.setState({ isAnimated: false})

  if (arg.message === "") {
    dialog.close()

  } else {
    var aRes = JSON.parse(arg.message);
    var aIndex = Number(aRes.index)
    var aValue = Number(aRes.value)
    var aColor = aRes.color

    if (!isNaN(aIndex) && Number.isInteger(aIndex)) {
      if (!isNaN(aValue)) {
        data[aIndex].value =  aValue
      }

      if (aColor) {
        data[aIndex].color =  aColor
      }

      chart.forceUpdate()
    }
  }
}

function openDialog() {
  chart.setState({ isAnimated: true})

  setTimeout(function() {
    localStorage.setItem("data", JSON.stringify(data))
    Office.context.ui.displayDialogAsync(
      kLocation + "/dialogPieChart.html",
      {height: 45, width: 25},

      function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
      }
    );
  }.bind(this), 800)
}

class Picker extends React.Component {
  constructor(props, context) {
    super(props, context);

    this.state = {
      color: props.color
    };
  }

  onColorClick = (color) => {
    this.setState({color: color.hex})
    this.props.onChange(color.hex, this.props.index)
  }

  render() {
    return <Popup trigger={<button style={{width: 16, height: 16, backgroundColor:this.state.color}}>    </button>}
                  position="bottom center">
             <div><CirclePicker color={this.state.color}
                                onChange={this.onColorClick}/>
             </div>
           </Popup>
  }
}

export default class PieChart extends React.Component {
  constructor(props, context) {
    super(props, context);

    this.state = {
      isAnimated: false,
    };
  }

  componentDidMount() {
    chart = this
  }

  generateData = (value, length = 5) =>
      d3.range(length).map((item, index) => ({
        date: index,
        value: value === null || value === undefined ? Math.random() * 100 : value
  }));

  onDataUpdate = async () => {
    openDialog()
  }

  onAdd = () => {
    data.push({date: data.length, value: (Math.floor(Math.random() * 100) + 10), color: aColors[Math.floor(Math.random() * aColors.length)]})
    this.forceUpdate()
  }

  onRemove = () => {
    data.pop()
    this.forceUpdate()
  }

  onColorChange = (color, index) => {
    data[index].color = color;
    this.forceUpdate()
  }

  parseData=()=>{
    let result = [];
    for (let i = 0; i < data.length; i++) {
      result.push([{width: 80,
                    overflow: 'clip',
                    value: data[i].value,},
                   {component: (<div style={{alignItems: 'center'}}>
                                  <Picker color={data[i].color} index={i} onChange={this.onColorChange}/>
                                </div>),
                    forceComponent: true,
                    width: 16}])
    }

    return result;
  }

  render() {
    return (
      <div className="container-screen">
        <div className="container-chart">
          <PieClass data={data}
                    width={200}
                    height={200}
                    innerRadius={60}
                    outerRadius={100}
          />
          <div className="chart-animation">
          <Lottie isStopped={!this.state.isAnimated}
              width={150}
              height={150}
              options={{ animationData: animation,
                         autoplay: false,
                         loop: true,
                       }}
      />
          </div>
          <br/>
          <Button className="chart-button"
                  buttonType={ButtonType.hero}
                  onClick={this.onDataUpdate}
          > EDIT </Button>
        </div>

        <div className="container-table">
          <ReactDataSheet style={{width: 100}}
                          data={this.parseData()}
                          valueRenderer={(cell) => cell.value}
                          onCellsChanged={changes => {
                            let aChanges = changes[0]
                            data[aChanges.row].value = aChanges.value
                            this.forceUpdate()}}
          />

          <Button buttonType={ButtonType.icon}
                  iconProps={{ iconName: "CirclePlus" }}
                  onClick={this.onAdd}
          />

          <Button buttonType={ButtonType.icon}
                  iconProps={{ iconName: "SkypeCircleMinus" }}
                  onClick={this.onRemove}
          />
        </div>
      </div>
    );
  }
}

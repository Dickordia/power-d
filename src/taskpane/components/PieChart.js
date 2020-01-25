/* global Button, console, Header, HeroList, HeroListItem, Office, Progress */

import * as React from "react";

import * as d3 from "d3";
import ReactDOM from "react-dom";

import PieClass from "./PieClass";
import { Button, ButtonType } from "office-ui-fabric-react";


var dialog = null;
var chart = null;
var data = [{date: 0, value: 36.473616873451826},
            {date: 1, value: 17.072667494331473},
            {date: 2, value: 54.911243181374736},
            {date: 3, value: 26.135479397225993},
            {date: 4, value: 76.9655970483639}]

function processMessage(arg) {
  if (arg.message === "") {
    dialog.close()

  } else {
    var aRes = JSON.parse(arg.message);
    var aIndex = Number(aRes.index)
    var aValue = Number(aRes.value)

    if (!isNaN(aIndex) && Number.isInteger(aIndex) && !isNaN(aValue)) {
      data[aIndex].value =  aValue
      chart.forceUpdate()
    }
  }
}

function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/dialogPieChart.html',
    {height: 45, width: 25},

    function (asyncResult) {
      dialog = asyncResult.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

export default class PieChart extends React.Component {
  constructor(props, context) {
    super(props, context);
  }

  componentWillMount() {
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

  render() {
    return (
      <div className="ms-welcome__header">
        <Button
         className="ms-welcome__action"
         buttonType={ButtonType.hero}
         onClick={this.onDataUpdate}
        > Update </Button>

        <PieClass
         data={data}
         width={200}
         height={200}
         innerRadius={60}
         outerRadius={100}
        />
      </div>
    );
  }
}

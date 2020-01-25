import * as React from "react";

import { OffCanvas, OffCanvasMenu, OffCanvasBody } from "react-offcanvas";

import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
/* global Button, console, Header, HeroList, HeroListItem, Office, Progress */

import { Slider } from 'office-ui-fabric-react/lib/Slider';

import USA from "@svg-maps/usa";
import { SVGMap } from "react-svg-map";
import "react-svg-map/lib/index.css";

import PieChart from "./PieChart"


export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);

    this.state = {
      listItems: [],
      value: 0,
      isPanel: false,
    };
  }

  click3() {
    this.setState({isPanel: !this.state.isPanel});
  }

  click2 = async () => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ },
        function (result) {
            if (result.status == "succeeded") {
                // If the getFileAsync call succeeded, then
                // result.value will return a valid File Object.
                var myFile = result.value;
                var sliceCount = myFile.sliceCount;
                var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
                app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

                // Get the file slices.
                getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
                app.showNotification("Error:", result.error.message);
            }
    });
  }

  click = async () => {
    Office.context.ui.displayDialogAsync(window.location.origin + "/dialog.html",
    { height: 50, width: 50, displayInIframe: true})
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div>
      <PieChart/>
      <div>
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.click}
        > React Web </Button>
      </div>



      <div className="ms-map">
        <SVGMap map={USA} />


      </div>

      <Slider
        label="S-L-I-D-E-R"
        min={0}
        max={50}
        step={10}
        defaultValue={20}
        showValue={true}
      />
      </div>
    );
  }
}

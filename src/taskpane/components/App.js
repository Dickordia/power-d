import * as React from "react";

import { OffCanvas, OffCanvasMenu, OffCanvasBody } from "react-offcanvas";
import * as d3 from "d3";
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
const kLocation = "https://localhost:3000"

import Lottie from 'lottie-react-web'
import animation from './anim4.json'
import './PieChart.css'

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);

    this.state = {
      listItems: [],
      value: 0,
      isAnimated: false,
    };
  }

  onShowReactWEB = () => {
    Office.context.ui.displayDialogAsync(
      kLocation + "/dialog.html",
    { height: 50, width: 50, displayInIframe: true})
  };

  render() {
    console.log(this.state.isAnimated);
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="container">
        <div className="header">
        <div>
          <Lottie
                  isStopped={!this.state.isAnimated}
                  width={70}
                  height={40}
                  options={{ animationData: animation,
                             autoplay: true,
                             loop: true,}}
          />
          </div>
        </div>
        <PieChart/>
      </div>
    );
  }
}

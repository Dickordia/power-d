'use strict';

import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";

const e = React.createElement;

class ProgressA extends React.Component {
  constructor(props) {
    super(props);
    this.state = { liked: false };
  }

  render() {
    if (this.state.liked) {
      return 'You liked this.';
    }

    return e(
      'button',
      { onClick: () => this.setState({ liked: true }) },
      'Like'
    );
  }
}


const domContainer = document.querySelector('#progress');
ReactDOM.render(<ProgressA/>, domContainer);

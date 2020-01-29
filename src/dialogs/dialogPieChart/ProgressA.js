'use strict';

import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";

class ProgressA extends React.Component {

  render() {
    return <div>
      <Button
        className="ms-welcome__action"
        buttonType={ButtonType.hero}
        iconProps={{ iconName: "ChevronRight" }}
        onClick={this.onShowReactWEB}
      > React Web CUSTOM </Button>
    </div>
  }
}

ReactDOM.render(<ProgressA/>, document.getElementById('progress'));

import React from "react";
import ReactDOM from "react-dom";
import Taskpane from "./Taskpane";

Office.onReady();

ReactDOM.render(<Taskpane />, document.getElementById("root"));

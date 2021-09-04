import * as React from "react";
import PropTypes from "prop-types";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import BarChart from "./chord";
import { renderToString } from "react-dom/server";

const dimensions = {
  width: 600,
  height: 300,
  margin: { top: 30, right: 30, bottom: 30, left: 60 },
};
/* global console, Excel */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
    this.myRef = React.createRef();
    this.markerRef = React.createRef();
  }

  columnToLetter(column) {
    var temp,
      letter = "";
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  letterToColumn(letter) {
    var column = 0,
      length = letter.length;
    for (var i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }

  buildAddressFromIndex(startCol, startRow, endCol, endRow) {
    var _startCol = this.columnToLetter(startCol + 1);
    var _endCol = this.columnToLetter(endCol + 1);
    var _startRow = (startRow + 1).toString();
    var _endRow = (endRow + 1).toString();
    var result = _startCol + _startRow + ":" + _endCol + _endRow;
    return result;
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        // const range = context.workbook.getSelectedRange();
        // var sheet = context.workbook.worksheets.getActiveWorksheet();

        // // Read the range address
        // range.load("address");
        // range.load("rowCount");
        // range.load("columnCount");
        // range.load("columnIndex");
        // range.load("rowIndex");

        // // Update the fill color
        // range.format.fill.color = "#f6dcb4";

        // await context.sync();
        // var elementCount = Math.max(range.rowCount, range.columnCount);
        // const headerRow = this.buildAddressFromIndex(
        //   range.columnIndex,
        //   range.rowIndex,
        //   range.columnIndex + elementCount - 1,
        //   range.rowIndex
        // );

        // const headerColumn = this.buildAddressFromIndex(
        //   range.columnIndex,
        //   range.rowIndex,
        //   range.columnIndex,
        //   range.rowIndex + elementCount - 1
        // );

        // const fillArea = this.buildAddressFromIndex(
        //   range.columnIndex + 1,
        //   range.rowIndex + 1,
        //   range.columnIndex + elementCount - 1,
        //   range.rowIndex + elementCount - 1
        // );

        // const dataArea = this.buildAddressFromIndex(
        //   range.columnIndex,
        //   range.rowIndex,
        //   range.columnIndex + elementCount - 1,
        //   range.rowIndex + elementCount - 1
        // );

        // var dataRange = sheet.getRange(headerRow);
        // dataRange.format.fill.color = "#b2ecea";
        // dataRange = sheet.getRange(headerColumn);
        // dataRange.format.fill.color = "#b2ecea";
        // dataRange = sheet.getRange(fillArea);
        // dataRange.format.fill.color = "#f6dcb4";
        // dataRange = sheet.getRange(dataArea);
        // dataRange.load("values, numberFormat");
        // //var dataValues = dataRange.values;
        // await context.sync();
        // var dataValues = dataRange.values;
        // const shapes = context.workbook.worksheets.getItem("Shapes").shapes;
        // const svgString = renderToString(<BarChart data={dataValues} />);
        // var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
        // line.name = "StraightLine";
        // // const sh = shapes.addSvg(svgString);
        // // sh.name = "test";
        // // sheet.shapes.addSvg(svgString);
        var shapes = context.workbook.worksheets.getItem("Sheet1").shapes;
        var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
        rectangle.left = 100;
        rectangle.top = 100;
        rectangle.height = 150;
        rectangle.width = 150;
        rectangle.fill.setSolidColor("green");
        rectangle.fill.transparency(0.5);
        //rectangle.fill = "#f6dcb4";
        rectangle.name = "Square";
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/fox_300.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

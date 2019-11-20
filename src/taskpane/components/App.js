import React, { Fragment } from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
/* global Button Header, HeroList, HeroListItem, Progress, Word */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      words: [],
      bulpitWords: [],
    };
  }

  componentDidMount() {
    this.highlight();
    console.log('test');
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  highlight = async () => {
    return Word.run(async context => {
      var body = context.document.body;
      console.log(body.paragraphs);
      context.load(body, 'text');
      const paragraph = context.document.body.insertParagraph("JOUOE", Word.InsertLocation.end);
      paragraph.font.color = "blue";
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();
      let text = [];
      paragraphs.items.forEach((item) => {
        let paragraph = item.text.trim();
        if (paragraph) {
          paragraph.split(" ").forEach((term) => {
            let currentTerm = term.trim();
            if (currentTerm) {
              text.push(currentTerm);
            }
          });
        }
      });
      this.setState({
        words: text,
        bulpitWords: [2,3],
      });
      console.log(text);
    });
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
        {/* <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" /> */}
        <main className="ms-welcome__main">
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.highlight}
          >
            Käivita
          </Button>
          {this.state.bulpitWords.length > 0 && (
            <p className="ms-font-l">
              Kantseliitsed sõnad:
            </p>
          )}
          {this.state.bulpitWords.map((bulpitIndex) => (
            <Fragment key={bulpitIndex}>
              <div>
                {this.state.words[bulpitIndex]}
              </div>
              <br />
            </Fragment>
          ))}
        </main>
      </div>
    );
  }
}

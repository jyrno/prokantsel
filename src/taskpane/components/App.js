import React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import BulpitWordItem from "./BulpitWordItem";
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
      context.document.body.select();

      await context.sync();
    });
  };

  highlight = async () => {
    return Word.run(async context => {

      var body = context.document.body;
      const myParagraphs = body.paragraphs;
      myParagraphs.load("text");
      await context.sync();

      myParagraphs.items.forEach(element => {
        console.log(element.text);
        element.font.color = "red";
      });
      // context.load(body, 'text');

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("JOUOE", Word.InsertLocation.end);
      paragraph.font.color = "blue";
      //console.log(paragraph);


      // await context.sync();
      //console.log("tes3322");

      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      //console.log("tes3322");

      await context.sync();

      let text = [];
      paragraphs.items.forEach((item) => {
        let paragraph = item.text.trim();
        console.log(item.text);

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
        bulpitWords: [
          {
            word: "Hello",
            description: "Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini Tee paremini ",
            type: "kantseliit",
          },
          {
            word: "Word",
            description: "Veel paremini",
            type: "paronyym",
          },
          {
            word: "Test",
            description: "Miks mitte veel paremini",
            type: "tarind",
          },
          {
            word: "Word",
            description: "Veel paremini",
            type: "paronyym",
          }
        ],
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
        <main className="ms-welcome__main taskpane">
          <Button
            className="ms-welcome__action bulpit__button"
            buttonType={ButtonType.hero}
            onClick={this.highlight}
          >
            Leia kantseliidid
          </Button>
          {this.state.bulpitWords.length > 0 && (
            <p className="ms-font-l">
              Kantseliitsed sõnad:
            </p>
          )}
          {this.state.bulpitWords.map((bulpitObject, idx) => (
            <BulpitWordItem
              key={idx}
              word={bulpitObject.word}
              description={bulpitObject.description}
              type={bulpitObject.type}
            />
          ))}
        </main>
      </div>
    );
  }
}

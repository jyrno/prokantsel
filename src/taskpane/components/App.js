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

  highlight = async () => {
    return Word.run(async context => {

      let documentParagraphs = context.document.body.paragraphs;
      documentParagraphs.load("text");
      await context.sync();
      // console.log(documentParagraphs.items[0]);
      const allParagraphs = [];
      const allSentencesObjects = [];
      const allSentences = [];

      documentParagraphs.items.forEach((paragraph) => {
        paragraph.load("text");
        allParagraphs.push(paragraph);
        const sentences = paragraph.split(["."], false /* trimDelimiters*/, true /* trimSpaces */);
        sentences.load("text");
        allSentencesObjects.push(sentences);
        
      });
      await context.sync();

      console.log(allParagraphs);
      console.log(allParagraphs.length);

      // console.log(allSentencesObjects);
      // console.log(allSentencesObjects.length);

      allSentencesObjects.forEach((sentencesObject) => {
        sentencesObject.items.forEach((sentence) => {
          sentence.load("text");
          allSentences.push(sentence);
        });
      });
      await context.sync();

      console.log(allSentences);
      console.log(allSentences.length);

      console.log(allSentences[0].getHtml());

      const words = [];
      const wordsObjects = [];
      allSentences.forEach((sentence) => {
        const words = sentence.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
        words.load("text");
        wordsObjects.push(words);
      });
      await context.sync();

      // console.log(wordsObjects);
      // console.log(wordsObjects.length);

      wordsObjects.forEach((wordObject) => {
        wordObject.items.forEach((word) => {
          word.load("text");
          words.push(word);
        });
      });

      await context.sync();
      console.log(words);
      console.log(words.length);


      this.setState({
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

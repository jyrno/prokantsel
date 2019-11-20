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

  highlight = async () => {
    return await Word.run(async context => {

      // let paragraph = context.document.body.paragraphs.getFirst();
      // let words = paragraph.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
      // words.load("text");
    
      // await context.sync();
      // console.log(words);


      let documentParagraphs = context.document.body.paragraphs;
      documentParagraphs.load("text");
      await context.sync();
      // console.log(documentParagraphs.items[0]);
      const allParagraphs = [];
      const allSentences = [];
      documentParagraphs.items.forEach(async (paragraph) => {
        paragraph.load("text");
        await context.sync();
        allParagraphs.push(paragraph);
      });

      console.log(allParagraphs);
      const allParagraphsCopy = [...allParagraphs];
      console.log(allParagraphs);
      console.log(allParagraphsCopy.slice(0).length);

      allParagraphsCopy.slice(0).forEach( async (paragraph) => {
        console.log(paragraph);

        const sentences = paragraph.split(["."], false /* trimDelimiters*/, true /* trimSpaces */);
        sentences.load("text");
        await context.sync();
        console.log(sentences);
        //console.log(sentences);
        //console.log(sentences.items.length);
        sentences.items.forEach(async (sentence) => {
            
            allSentences.push(sentence);
            sentence.load("text");
            await context.sync();
      //     let words = sentence.split([" "], false /* trimDelimiters*/, true /* trimSpaces */);

        });
      });


      // let sentences = documentParagraphs.split(["."], false /* trimDelimiters*/, true /* trimSpaces */);
      // sentences.load("text");
      // await context.sync();
      // console.log(sentences);


      // let documentParagraphs = context.document.body.paragraphs;
      // documentParagraphs.items.forEach(async (paragraph) => {
      //   console.log(paragraph);
      //   let sentences = paragraph.split(["."], false /* trimDelimiters*/, true /* trimSpaces */);
      //   sentences.load("text");
      //   await context.sync();
      //   console.log(sentences);
      //   sentences.items.forEach(async (sentence) => {
      //     let words = sentence.split([" "], false /* trimDelimiters*/, true /* trimSpaces */);
      //     words.load("text");
      //     await context.sync();
      //     console.log(words);
      //   });
      // });
      //let sentences = documentParagraphs.split(["."], false /* trimDelimiters*/, true /* trimSpaces */);

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

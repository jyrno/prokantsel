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
          if(word.text == "dfbdf")
          {
            word.font.color = "yellow";
          }
          words.push(word);
        });
      });

      await context.sync();
      console.log(words);
      console.log(words.length);

      words.forEach((word) => {
        if(word.text == "dfbdf")
        {
          word.font.bold = true;
        }
      });


      console.log(words[0].text);
      console.log(words[0].getHtml());
      console.log(words[0].getOoxml());

      var html = words[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });

      words[0].insertHtml(`<HTML>
      <HEAD></HEAD>
      <BODY>
      <div class="OutlineGroup"><div class="OutlineElement Ltr"><div class="ParaWrappingDiv"><p class="Paragraph" xml:lang="EN-US" lang="EN-US" paraid="0" paraeid="{33b92d5c-2cb1-4b2e-9e40-76823e625498}{240}" style="font-weight: normal; font-style: normal; vertical-align: baseline; font-family: &quot;Segoe UI&quot;, Tahoma, Verdana, Sans-Serif; background-color: transparent; color: windowtext; text-align: left; margin: 0px 0px 10.6667px; padding-left: 0px; padding-right: 0px; text-indent: 0px; font-size: 6pt;"><span data-contrast="none" class="TextRun" xml:lang="EN-US" lang="EN-US" style="color: rgb(255, 0, 0); background-color: transparent; font-size: 11pt; font-family: Calibri, Calibri_MSFontService, sans-serif; font-kerning: none; line-height: 19.425px;"><span class="NormalTextRun" style="background-color: blue;">idsdsdsdto</span></span><span class="EOP" style="font-size: 11pt; line-height: 19.425px; font-family: Calibri, Calibri_MSFontService, sans-serif;">&nbsp;</span></p></div></div></div><span class="WACImageGroupContainer"></span><span data-contrast="none" class="TextRun" xml:lang="EN-US" lang="EN-US" style="color: rgb(255, 0, 0); background-color: transparent; font-size: 11pt; font-family: Calibri, Calibri_MSFontService, sans-serif; font-kerning: none; line-height: 19.425px;"></span><span class="NormalTextRun" style="background-color: inherit;"></span>
      </BODY>
      </HTML>`, Word.InsertLocation.replace);

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
      
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo))}});
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

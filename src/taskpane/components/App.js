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

      let documentBody = context.document.body;
      documentBody.load("text");
      await context.sync();

      let documentParagraphs = context.document.body.paragraphs;
      documentParagraphs.load("text");
      await context.sync();

      const response = await fetch("https://demo2624123.mockable.io/", {
          method: 'POST', // *GET, POST, PUT, DELETE, etc.
          mode: 'cors', // no-cors, *cors, same-origin
          cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
          credentials: 'same-origin', // include, *same-origin, omit
          headers: {
            'Content-Type': 'application/json'
            // 'Content-Type': 'application/x-www-form-urlencoded',
          },
          redirect: 'follow', // manual, *follow, error
          referrer: 'no-referrer', // no-referrer, *client
          body: JSON.stringify(documentBody.text) 
      });

      const responseJson = await response.json();
      const matchingStrings = [];
      const searchResultObjects = [];

      responseJson.analysis.forEach(async analysis => {
        console.log("Searching for word: " + analysis.text);
        const searchResult = documentBody.search(analysis.text);
        searchResult.load("text");
        searchResultObjects.push(searchResult);
        console.log(searchResult);
      });
      await context.sync();  
      console.log(searchResultObjects);

      searchResultObjects.forEach(object => {
        object.items.forEach(resultItem => {
          resultItem.load("text");
          matchingStrings.push(resultItem);
        });
      });
      await context.sync();
      console.log(matchingStrings);

      matchingStrings.forEach(item => {
        item.load("text");
        item.font.color = 'purple';
        item.font.highlightColor = 'pink';
        item.font.bold = true;
      });
      await context.sync();
      
      const complexityValue = responseJson.complexity.coef;
      const complexityDescription = responseJson.complexity.text;
      console.log(complexityValue);
      console.log(complexityDescription);
      

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

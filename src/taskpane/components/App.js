import React from "react";
import { Button, ButtonType, MarqueeSelection } from "office-ui-fabric-react";
import BulpitWordItem from "./BulpitWordItem";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { isObject } from "util";
/* global Button Header, HeroList, HeroListItem, Progress, Word */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      bulpitWords: [],
      searchResults: [],
    };
    this.parseResponse = this.parseResponse.bind(this);
  }

  componentDidMount() {
    console.log("OLEN Mountis");

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

  componentWillUnmount() {
    console.log("OLEN UNMOUNTIS");
    this.cleanDocument();
}

  parseResponse = (responseJson) => {
    this.setState({
      bulpitWords: responseJson.analysis,
      complexity: responseJson.complexity,
    });
  };

  highlight = async () => {
    return Word.run(async context => {
      this.cleanDocument();
      let documentBody = context.document.body;
      documentBody.load("text");
      await context.sync();

      let documentParagraphs = context.document.body.paragraphs;
      documentParagraphs.load("text");
      await context.sync();
      console.log("Update");

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
          body: {"text": documentBody.text}
      });
      console.log("Update");
      const responseJson = await response.json();

      // const response2 = await fetch("https://172.31.99.247:5000/")

      responseJson.analysis[2] = {"text": "politseiniku poolt", "type": "POOLT_TARIND"};
      responseJson.analysis[3] = {"text": "politseiniku poolt", "type": "POOLT_TARIND"};

      console.log(responseJson);
      this.parseResponse(responseJson);
      const matchingStrings = [];
      const searchResultObjects = [];
      const searchedVerbs = [];

      responseJson.analysis.forEach(async analysis => {
        if(searchedVerbs.indexOf(analysis.text) == -1)
        {
          console.log("Searching for word: " + analysis.text);
          const searchResult = documentBody.search(analysis.text);
          searchedVerbs.push(analysis.text);
          searchResult.load("text");
          searchResultObjects.push(searchResult);
          console.log(searchResult);
        }
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
      console.log("Not state list");
      console.log(matchingStrings);
      this.setState(state => ({
        searchResults: matchingStrings,
      }));
      // console.log("State list");
      // console.log(searchResults);

      matchingStrings.forEach(item => {
        item.load("text");
        item.font.color = 'purple';
        item.font.highlightColor = 'pink';
        //item.font.bold = true;
      });
      await context.sync();

      const complexityValue = responseJson.complexity.coef;
      const complexityDescription = responseJson.complexity.text;
      console.log(complexityValue);
      console.log(complexityDescription);
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo))}});
  };

  replaceSinglePhrase = (phraseObject, phraseToReplaceWith) => {
    phraseObject.insertText(phraseToReplaceWith, Word.InsertLocation.replace);
  };

  cleanSignlePhrase = async (phraseObject) => {
    phraseObject.font.highlightColor = 'white';
    phraseObject.font.color = 'black';
    await context.sync();
  };

  cleanDocument = async () => {
    return Word.run(async context => {
      let documentBody = context.document.body;
      documentBody.load("text");
      documentBody.font.highlightColor = 'white';
      documentBody.font.color = 'black';
      await context.sync();

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
          {this.state.complexity && (
            <p><b>Analüüsi tulemus: </b>{this.state.complexity.text}</p>
          )}
          {this.state.bulpitWords.map((bulpitObject, idx) => (
            <BulpitWordItem
              key={idx}
              word={bulpitObject.text}
              type={bulpitObject.type}
              verb={bulpitObject.verb}
              searchObjects={this.state.searchResults}
              onIgnore={this.cleanSignlePhrase}
              onReplace={this.replaceSinglePhrase}
            />
          ))}
        </main>
      </div>
    );
  }
}

import React from "react";
import { Button, ButtonType, MarqueeSelection } from "office-ui-fabric-react";
import BulpitWordItem from "./BulpitWordItem";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { isObject } from "util";
/* global Button Header, HeroList, HeroListItem, Progress, Word */

const colors = {
  'PARONUUM': '#ffcfe7',
  'NOMINALISATSIOON': '#efcfff',
  'POOLT_TARIND': '#f5f783',
  'OLEMA_KESKSONA': '#d0f5ef',
  'KANTSELIIT': '#c5d2fc',
  'LIIGNE_MITMUS': '#d3f5d4',
  'SAAV_KAANE': '#fcebca',
  'LT_MAARSONA': '#f5d5d5'
}

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
    this.cleanDocument();
  }

  async componentWillUnmount() {
    await this.cleanDocument();
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
      // console.log("Update");

      const response = await fetch("https://172.31.98.184:5000/check", {
          method: 'POST', // *GET, POST, PUT, DELETE, etc.
          mode: 'cors', // no-cors, *cors, same-origin
          cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
          credentials: 'same-origin', // include, *same-origin, omit
          // 'Content-Type': 'application/json',
          headers: {
            'Content-Type': 'application/json'
            // 'Content-Type': 'application/x-www-form-urlencoded',
          },
          redirect: 'follow', // manual, *follow, error
          referrer: 'no-referrer', // no-referrer, *client
          body: JSON.stringify({"text": documentBody.text})
      });
      // console.log("Update");
      const responseJson = await response.json();

      // console.log(responseJson);
      this.parseResponse(responseJson);
      const matchingStrings = [];
      const searchResultObjects = [];
      const searchedVerbs = [];

      responseJson.analysis.forEach(async analysis => {
        if(searchedVerbs.indexOf(analysis.text) == -1)
        {
          // console.log("Searching for word: " + analysis.text);
          const searchResult = documentBody.search(analysis.text);
          searchedVerbs.push(analysis.text);
          searchResult.load("text");
          // console.log('found object');
          // console.log(searchResult);
          searchResultObjects.push(searchResult);
          // console.log(searchResult);
        }
      });
      await context.sync();
      // console.log(searchResultObjects);
      //
      // console.log('searchResults:');
      // console.log(searchResultObjects);
      searchResultObjects.forEach(object => {
        object.items.forEach(resultItem => {
          resultItem.load("text");
          matchingStrings.push(resultItem);
        });
      });
      await context.sync();
      // console.log("Not state list");
      // console.log(matchingStrings);
      this.setState(state => ({
        searchResults: matchingStrings,
      }));

      // console.log('doing coloring');
      // console.log(matchingStrings);
      matchingStrings.forEach(item => {
        item.load("text");
        item.font.highlightColor = '#c5d2fc';
        // console.log('item');
        // console.log(item);
        // console.log('responseJson');
        // console.log(responseJson.analysis);
        // console.log(responseJson.analysis);
        // console.log(item.text);
        const correctItem = responseJson.analysis.find((analysedItem) => analysedItem.text.toLowerCase() === item.text.toLowerCase());
        // console.log(correctItem);
        // console.log(correctItem.type);
        // const itemType = responseJson.analysis.find((analysedItem) => analysedItem.text === item.text).type;
        // console.log('item text:');
        // console.log(item.text);
        // console.log('itemType');
        // console.log(itemType);
        // console.log('color:');
        // console.log(colors);
        // console.log(colors[itemType]);
        if (correctItem) {
          item.font.highlightColor = colors[correctItem.type];
        }
      });
      await context.sync();

      const complexityValue = responseJson.complexity.coef;
      const complexityDescription = responseJson.complexity.text;
      // console.log(complexityValue);
      // console.log(complexityDescription);
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo))}});
  };

  replaceSinglePhrase = (phraseObject, phraseToReplaceWith) => {
    return Word.run(async context => {
      let documentBody = context.document.body;
      documentBody.load("text");
      let searchRes = documentBody.search(phraseObject.text);
      searchRes.load("text");
      await context.sync();
      searchRes.items[0].insertText(phraseToReplaceWith, Word.InsertLocation.replace);
      await context.sync();
      this.setState(state => ({
        bulpitWords: state.bulpitWords.filter((bulpit) => bulpit.text.toLowerCase() !== phraseObject.text.toLowerCase()),
      }));
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo))}});
  };

  cleanSignlePhrase = async (phraseObject) => {
    // console.log('cleanSignlePhrase');
    // console.log(phraseObject);
    return Word.run(async context => {
    let documentBody = context.document.body;
    documentBody.load("text");
    let searchRes = documentBody.search(phraseObject.text.toLowerCase());
    searchRes.load("text");
    await context.sync();
    searchRes.items[0].font.highlightColor = 'white';
    searchRes.items[0].font.color = 'black';
    await context.sync();
    // console.log('cleaning');
    // console.log(state.bulpitWords);
    // console.log(phraseObject.text.toLowerCase());
    // console.log(state.bulpitWords.filter((bulpit) => bulpit.text.toLowerCase() !== phraseObject.text.toLowerCase()));
    this.setState(state => ({
      bulpitWords: state.bulpitWords.filter((bulpit) => bulpit.text.toLowerCase() !== phraseObject.text.toLowerCase()),
    }));
  })
  .catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo))}});
  };

  cleanDocument = async () => {
    return Word.run(async context => {
      let documentBody = context.document.body;
      documentBody.load("text");
      documentBody.font.highlightColor = 'white';
      // documentBody.font.color = 'black';
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
            Analüüsi
          </Button>
          {this.state.complexity && (
            <p className="bulpit__complexity">
              {this.state.complexity.coef > 75 ? (
                <svg className="bulpit__complexity-img bulpit__complexity-img--bad" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm.001 14c-2.332 0-4.145 1.636-5.093 2.797l.471.58c1.286-.819 2.732-1.308 4.622-1.308s3.336.489 4.622 1.308l.471-.58c-.948-1.161-2.761-2.797-5.093-2.797zm-3.501-6c-.828 0-1.5.671-1.5 1.5s.672 1.5 1.5 1.5 1.5-.671 1.5-1.5-.672-1.5-1.5-1.5zm7 0c-.828 0-1.5.671-1.5 1.5s.672 1.5 1.5 1.5 1.5-.671 1.5-1.5-.672-1.5-1.5-1.5z"/></svg>
              ) : this.state.complexity.coef > 50 ? (
                <svg className="bulpit__complexity-img bulpit__complexity-img--ok" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm4 17h-8v-2h8v2zm-7.5-9c-.828 0-1.5.671-1.5 1.5s.672 1.5 1.5 1.5 1.5-.671 1.5-1.5-.672-1.5-1.5-1.5zm7 0c-.828 0-1.5.671-1.5 1.5s.672 1.5 1.5 1.5 1.5-.671 1.5-1.5-.672-1.5-1.5-1.5z"/></svg>
              ) : (
                <svg className="bulpit__complexity-img bulpit__complexity-img--good" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm6 14h-12c.331 1.465 2.827 4 6.001 4 3.134 0 5.666-2.521 5.999-4zm0-3.998l-.755.506s-.503-.948-1.746-.948c-1.207 0-1.745.948-1.745.948l-.754-.506c.281-.748 1.205-2.002 2.499-2.002 1.295 0 2.218 1.254 2.501 2.002zm-7 0l-.755.506s-.503-.948-1.746-.948c-1.207 0-1.745.948-1.745.948l-.754-.506c.281-.748 1.205-2.002 2.499-2.002 1.295 0 2.218 1.254 2.501 2.002z"/></svg>
              )}
              {this.state.complexity.text}
            </p>
          )}
          {this.state.bulpitWords && this.state.bulpitWords.map((bulpitObject) => (
            <BulpitWordItem
              key={Math.random()}
              word={bulpitObject.text}
              type={bulpitObject.type}
              verb={bulpitObject.verb}
              synonyms={bulpitObject.synonyms}
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

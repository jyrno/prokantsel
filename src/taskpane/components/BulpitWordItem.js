import React, { Component } from "react";
import { descriptions } from "../../../helpers/index.js";
import { collapseTextChangeRangesAcrossMultipleVersions } from "typescript";

export default class BulpitWordItem extends Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isOpen: false,
      isHidden: false,
    };
    this.handleItemClick = this.handleItemClick.bind(this);
    this.handleIgnore = this.handleIgnore.bind(this);
    this.handleReplace = this.handleReplace.bind(this);
  }

  handleItemClick = () => {
    this.setState(state => ({
      isOpen: !state.isOpen
    }));
    const object = this.props.searchObjects.find((searchObject) => searchObject.text === this.props.word);
  }

  handleIgnore = () => {
    // console.log('in handleIgnore');
    this.setState({
      isHidden: true,
      isOpen: false,
    });
    // console.log('handleIgnore');
    setTimeout(() => {
      const phraseObject = this.props.searchObjects.find((searchObject) => searchObject.text.toLowerCase() === this.props.word.toLowerCase());
      this.props.onIgnore(phraseObject);
      // console.log('triggering onIgrnore');
    }, 500);
  }

  handleReplace = (synonym) => {
    // console.log('in handleReplace');
    // console.log(synonym);
    this.setState({
      isHidden: true,
      isOpen: false,
    });
    setTimeout(() => {
      const phraseObject = this.props.searchObjects.find((searchObject) => searchObject.text === this.props.word);
      this.props.onReplace(phraseObject, synonym);
    }, 500);
  }

  render() {
    const { word, type, verb, synonyms } = this.props;
    // console.log(type);
    // console.log(descriptions[type]);

    return (
      <div
        className={"bulpit bulpit--" + (this.state.isOpen ? 'open' : this.state.isHidden ? 'hidden' : 'closed')}
        ref={(c) => { this.bulpitWordItem = c; }}
        style={{
          maxHeight: this.state.isHidden ? 0 : this.state.isOpen ? this.bulpitWordItem.scrollHeight : 36,
          marginBottom: this.state.isHidden ? 0 : 8,
          boxShadow: this.state.isHidden ? '0px 1px 5px 0px rgba(0,0,0,0)' : '0px 1px 5px 0px rgba(0,0,0,.2)',
          transitionDuration: this.state.isHidden ? '.4s' : '.2s',
        }}
      >
        <div className="bulpit__word-wrapper">
          <div className="bulpit__container">
            <span className={"bulpit__indicator bulpit__indicator--" + type}></span>
            <p className="bulpit__word">
              {word}
            </p>
          </div>
          <div className="bulpit__container">
            <div className="bulpit__ignore" onClick={this.handleIgnore}>
              <span>Ignoreeri</span>
            </div>
            <div className="bulpit__arrow-wrapper" onClick={this.handleItemClick}>
              <i className="bulpit__arrow"></i>
            </div>
          </div>
        </div>
        <p className="bulpit__message">
          {type === 'NOMINALISATSIOON' ? (
            descriptions[type].description[verb].description
          ) : (
            descriptions[type].description
          )}
        </p>
        <div className="bulpit__replace">
          {synonyms && <span className="bulpit__replace-title">Asenda:</span>}
          {synonyms && synonyms.map((synonym, idx) => (
            <span className="bulpit__synonym" key={idx} onClick={() => this.handleReplace(synonym)}>{synonym}</span>
          ))}
        </div>
        <div className="bulpit__examples">
          <span className="bulpit__example-title">
            Näide:
          </span>
          <p className="bulpit__example--wrong">
            {type === 'NOMINALISATSIOON' ? (
              descriptions[type].description[verb].example_wrong
            ) : (
              descriptions[type].example_wrong
            )}
          </p>
          <p className="bulpit__example--correct">
            {type === 'NOMINALISATSIOON' ? (
              descriptions[type].description[verb].example_correct
            ) : (
              descriptions[type].example_correct
            )}
          </p>
        </div>
      </div>
    );
  }
}

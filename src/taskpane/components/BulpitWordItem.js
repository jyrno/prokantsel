import React, { Component } from "react";
import { descriptions } from "../../../helpers/index.js";

export default class BulpitWordItem extends Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isOpen: false,
    };
    this.handleItemClick = this.handleItemClick.bind(this);
  }

  handleItemClick = () => {
    console.log('click');
    this.setState(state => ({
      isOpen: !state.isOpen
    }));
    console.log(this.bulpitWordItem.scrollHeight);
  }

  render() {
    const { word, type, verb } = this.props;

    return (
      <div
        className={"bulpit bulpit--" + (this.state.isOpen ? 'open' : 'hidden')}
        ref={(c) => { this.bulpitWordItem = c; }}
        style={{ maxHeight: this.state.isOpen ? this.bulpitWordItem.scrollHeight : 20 }}
      >
        <div className="bulpit__word-wrapper">
          <div className="bulpit__container">
            <span className={"bulpit__indicator bulpit__indicator--" + type}></span>
            <p className="bulpit__word">
              {word}
            </p>
          </div>
          <div className="bulpit__container">
            <div className="bulpit__ignore">
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

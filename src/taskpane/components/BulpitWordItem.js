import React, { Component } from "react";

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
    const { word, description, type } = this.props;

    return (
      <div
        className={"bulpit bulpit--" + (this.state.isOpen ? 'open' : 'hidden')}
        onClick={this.handleItemClick}
        ref={(c) => { this.bulpitWordItem = c; }}
        style={{ maxHeight: this.state.isOpen ? this.bulpitWordItem.scrollHeight : 20 }}
      >
        <div className="bulpit__word-wrapper">
          <span className={"bulpit__indicator bulpit__indicator--" + type}></span>
          <p className="bulpit__word">
            {word}
          </p>
        </div>
        <p className="bulpit__message">
          {description}
        </p>
      </div>
    );
  }
}

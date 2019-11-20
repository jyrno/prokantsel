import * as React from "react";

export default class HeroList extends React.Component {
  render() {
    const { children } = this.props;

    return (
      <main className="ms-welcome__main">
        {children}
      </main>
    );
  }
}

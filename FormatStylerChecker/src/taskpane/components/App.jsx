import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import DocumentCheck from "./DocumentCheck"
import TextInsertion from "./TextInsertion";
import { makeStyles, Button } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText, checkDocument } from "../taskpane";
import { fixBasicFormatting } from "./fixers";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      {/* <TextInsertion insertText={insertText} /> */}
      <DocumentCheck checkDocument={checkDocument}/>
   
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', marginTop: '20px' }}>
        <h3>Automated Corrections</h3>
        <Button appearance="primary" onClick={fixBasicFormatting}>
    Fix Basic Formatting
</Button>
      </div>
      </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

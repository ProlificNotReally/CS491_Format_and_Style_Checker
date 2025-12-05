import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import PropTypes from "prop-types";

const useStyles = makeStyles({
//   instructions: {
//     fontWeight: tokens.fontWeightSemibold,
//     marginTop: "20px",
//     marginBottom: "10px",
//   },
//   textPromptAndInsertion: {
//     display: "flex",
//     flexDirection: "column",
//     alignItems: "center",
//   },
//   textAreaField: {
//     marginLeft: "20px",
//     marginTop: "30px",
//     marginBottom: "20px",
//     marginRight: "20px",
//     maxWidth: "50%",
//   },
    checkdoc: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
    },
});

const CheckDocument = (props) => {
    const styles = useStyles();
    
    const handleDocumentCheck = async () =>{
        await props.checkDocument();
    }

    return (
        <div className={styles.checkdoc}>
            <Button appearance="primary" disabled={false} size="large" onClick={handleDocumentCheck}>Check Document</Button>
        </div>
    );
}

CheckDocument.propTypes = {
    checkDocument : PropTypes.func.isRequired,
};

export default CheckDocument;

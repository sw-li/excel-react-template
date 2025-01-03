import * as React from "react";
import PropTypes from "prop-types";
import { makeStyles } from "@fluentui/react-components";
import { HashRouter, Route, Switch, Link } from "react-router-dom";
import Box from "@mui/material/Box";
import Tabs from "@mui/material/Tabs";
import Tab from "@mui/material/Tab";
import AboutScreen from "./screens/AboutScreen";
import UsersScreen from "./screens/UsersScreen";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  tabsContainer: {
    overflowX: "auto",
  },
});

const App = () => {
  const [value, setValue] = React.useState(0);
  const styles = useStyles();
  const handleChange = (event, newValue) => {
    setValue(newValue);
  };

  return (
    <div className={styles.root}>
      <HashRouter>
        <div>
          <Box sx={{ width: "100%" }}>
            <div className={styles.tabsContainer}>
              <Tabs value={value} onChange={handleChange} aria-label="nav tabs example" role="navigation">
                <Tab label="About" to="/about" component={Link} />
                <Tab label="Users" to="/users" component={Link} />
              </Tabs>
            </div>
          </Box>

          <Switch>
            <Route path="/">
              <AboutScreen />
            </Route>
            <Route path="/users">
              <UsersScreen />
            </Route>
          </Switch>
        </div>
      </HashRouter>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

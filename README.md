
# Excel Taskpane Add-in Template with React and Tabs

Welcome to the **Excel Taskpane Add-in Template** repository! This project provides a starting point for building an Excel Taskpane add-in using **React**, with a pre-configured setup for tabs. It’s designed to help developers quickly build modern, user-friendly Excel add-ins with a focus on modularity and scalability.

---

## Features

- **React Framework**: Leverages the power of React for building dynamic and responsive user interfaces.
- **Tab Navigation**: Preconfigured tab structure for easy navigation within the taskpane.
- **Excel Add-in Ready**: Fully compliant with Microsoft Office Add-in development guidelines.
- **Customizable**: Easily extend and customize components for your specific needs.

---

## Prerequisites

Before you get started, ensure you have the following installed on your development machine:

1. **Node.js** (LTS recommended)
2. **npm** or **yarn**
3. **Office Add-in prerequisites**:
   - Office 365 or Microsoft Office 2016+
   - [Office Add-in prerequisites](https://learn.microsoft.com/en-us/office/dev/add-ins/get-started/)

---

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/your-username/excel-taskpane-addin-template.git
cd excel-taskpane-addin-template
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Run the Project Locally

Start the development server:

```bash
npm start
```

This will launch Excel and a local development server, which uses the manifest file in root folder to sideload the add-in into Excel 


## Project Structure

```
|-- .babelrc
|-- .eslintrc.json
|-- .gitignore
|-- .vscode/
|   |-- extensions.json
|   |-- launch.json
|   |-- settings.json
|   |-- tasks.json
|-- assets/
|   |-- icon-128.png
|   |-- icon-16.png
|   |-- icon-32.png
|   |-- icon-64.png
|   |-- icon-80.png
|   |-- logo-filled.png
|-- babel.config.json
|-- manifest.xml
|-- package-lock.json
|-- package.json
|-- src/
|   |-- commands/
|   |   |-- commands.html
|   |   |-- commands.js
|   |-- functions/
|   |-- taskpane/
|   |   |-- App.jsx
|   |   |-- index.jsx
|   |   |-- screens/
|   |   |   |-- AboutScreen.jsx
|   |   |   |-- UsersScreen.jsx
|   |   |-- taskpane.html
|-- tsconfig.json
|-- webpack.config.js
```



## Customization

### Adding a New Tab

1. Create a new React screen component in the `src/taskpane/screens` directory.
2. Import and include the component in the `App.jsx` file.
3. Update the tab configuration with the new tab name and reference.
   
```bash
return (
    <div className={styles.root}>
      <HashRouter>
        <div>
          <Box sx={{ width: "100%" }}>
            <div className={styles.tabsContainer}>
              <Tabs value={value} onChange={handleChange} aria-label="nav tabs example" role="navigation">
                <Tab label="About" to="/about" component={Link} />
                <Tab label="Users" to="/users" component={Link} />
                <Tab label="NewTab" to="/newTab" component={Link} />
              </Tabs>
            </div>
          </Box>
          <Switch>
            <Route path="/">
              <AboutScreen />
            </Route>
            <Route path="/newTab">
              <NewTabScreen />
            </Route>
            <Route path="/users">
              <UsersScreen />
            </Route>
          </Switch>
        </div>
      </HashRouter>
    </div>
  );
```

### Modifying Styles

The template included MUI package. Feel free to use customized styling methods. 

---

## Contributing

Contributions are welcome! If you’d like to contribute, please:

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature-name`).
3. Commit your changes (`git commit -m "Add feature"`).
4. Push the branch (`git push origin feature-name`).
5. Open a pull request.

---

## License

This project is licensed under the MIT License. See the MIT License file for more details.

---

## Acknowledgements

- [React](https://reactjs.org/)
- [Office JavaScript API](https://learn.microsoft.com/en-us/javascript/api/overview/excel)

The basic strucutre is initiated using generator-office, then enriched with the React.
Feel free to adapt and expand this template as per your project requirements!

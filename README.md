# learning-word-add-in-take-three
Developing a Microsoft Word Add-in using JavaScript and Node.js. 

## Use case

Construction contracts generally have a list of defined terms (a dictionary) that is used throughout their general conditions to avoid any ambiguity in the interpretation of its terms. 

A risk arises where contract drafters copy and paste clauses from other contracts into a new one - these clauses may contain Capitalised Terms which are not defined in the new contract, causing ambiguity in its interpretation. 

I intend to have the add-in check through the body of the contract for any capitalised terms that do not have a corresponding defined term in the definitions section of a construction contract.

Foreseeable issues:
- defined terms that begin a sentence may be hard to identify
- how to integrate with a company's existing tech

Another use case may be to check for cross referencing errors without having to update fields - I will park this idea for now.

## Links

Original code template: https://github.com/OfficeDev/Office-Addin-TaskPane-JS 

Followed this tutorial: https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial#insert-a-range-of-text

Useful overview of how Office Add-ins work: https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins 

## Fluent UI framework 

For front end work, use the Fluent UI framework:
npm install @fluentui/react
then see docs: https://developer.microsoft.com/en-us/fluentui/#/controls/web

Example of how to import component (and pass a prop):
import { DefaultButton } from 'office-ui-fabric-react';

const MyComponent = () => {
  return (
    <div>
      <DefaultButton iconProps={{ iconName: 'Mail' }}>Send Mail</DefaultButton>
    </div>
  );
};

Some Fabric components take in a render functions to allow customizing certain parts of the component. An example with TextField:
import { TextField } from 'office-ui-fabric-react';

const MyComponent = () => {
  return (
    <div>
      <TextField onRenderPrefix={() => <Icon iconName="Search" />} />
      <TextField onRenderPrefix={() => 'hello world'} />
    </div>
  );
};

## To set up a project:
1. ensure yeoman is installed globally: npm install -g yo generator-office
2. to create a new project: yo office
3. to test: npm start
4. Before commiting your changes to git, create a .gitignore with the following code: node_modules/

I failed many times to get the dev server running right, espcially where you create more than one project (this took me three days and two deleted repos to debug). Some tips:
- clear your Word cache (go to /Users/<username>/Library/Containers/com.microsoft.Word/Data and clear the contents to ensure no other manifest.xml is running)
- ensure your dev server is not running another project's files (I could not get localserver:3000 to stop running the files of the first project I created so I amended the manifest.xml and package.json to run localhost:3001 and it worked - hopefully this error wont occur when pushing to prod)

### issue logs from failed attempts
1. The wrong manifest.xml file is being run (the application name used is from the old apps made)
2. The wrong directory is being used for 'loader' files in the node packages (the path is shown in the Word error: "Html Webpack Plugin: Error: Child compilation failed: Module not found:)
3. the webpack.config.js file seems to be doing something wrong as the error message refers to webpack
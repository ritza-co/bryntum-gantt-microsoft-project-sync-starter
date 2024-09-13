# How to connect and sync Bryntum Gantt to Microsoft Project

The code for the complete app is on the `completed-gantt` branch.

## Getting started

The starter repository uses [Vite](https://vitejs.dev/), which is a development server and JavaScript bundler. You’ll need Node.js version 18+ for Vite to work. 
Install the Vite dev dependency by running the following command: 

```sh
npm install
```

Install the Bryntum Gantt component by following [step 1](https://bryntum.com/products/gantt/docs/guide/Gantt/quick-start/javascript-npm#access-to-npm-registry) and [step 4](https://bryntum.com/products/gantt/docs/guide/Gantt/quick-start/javascript-npm#install-component) of the [vanilla JavaScript with npm setup guide](https://bryntum.com/products/gantt/docs/guide/Gantt/quick-start/javascript-npm).

## Running the app

Run the local dev server using the following command:

```sh
npm run dev
```

You'll see a Bryntum Gantt with 2 tasks and a dependency between the tasks:

![Initial Bryntum Gantt with two tasks and a dependency between the tasks](images/initial-app.png)
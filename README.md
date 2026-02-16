# Donation App - People Management System


A React Native Web application for managing people with their details, using MongoDB as the database.

## Features

- View list of people in a table format
- Add new people with all required details
- Edit existing people via popup modal
- Delete people from the list
- Clean and simple UI design

## Prerequisites

Before you begin, make sure you have the following installed on your PC:

1. **Node.js** (v14 or higher, v20.x.x recommended) - [Download here](https://nodejs.org/)
   - Includes npm (Node Package Manager)
   - Tested with Node.js v20.19.6 ✅
   - Verify: `node --version` and `npm --version`

2. **MongoDB** (Community Server) - [Download here](https://www.mongodb.com/try/download/community)
   - Must be installed locally
   - Should run on default port 27017
   - Verify: `mongod --version`
   - Make sure MongoDB service is running (check Windows Services)

3. **Modern Web Browser** (Chrome, Edge, Firefox, or Safari)

> 📖 **For detailed installation instructions, see [DEPENDENCIES.md](DEPENDENCIES.md)**

## Installation

1. Install MongoDB locally if you haven't already:
   - Download from: https://www.mongodb.com/try/download/community
   - Follow installation instructions for your OS
   - Make sure MongoDB service is running

2. Create the backend `.env` file:
   ```bash
   cd backend
   copy env.example .env
   ```
   (On Linux/Mac: `cp env.example .env`)

3. Install all dependencies:
   ```bash
   npm run install-all
   ```

   Or install manually:
   ```bash
   npm install
   cd backend && npm install
   cd ../frontend && npm install
   ```

## Running the Application

1. Make sure MongoDB is running on your local machine:
   ```bash
   # On Windows (if MongoDB is installed as a service, it should start automatically)
   # Or start manually:
   mongod
   ```

2. Start both backend and frontend servers:
   ```bash
   npm run dev
   ```

   Or start them separately:
   ```bash
   # Terminal 1 - Backend
   npm run server

   # Terminal 2 - Frontend
   npm run client
   ```

3. Open your browser and navigate to:
   ```
   http://localhost:3000
   ```

## API Endpoints

- `GET /api/persons` - Get all people
- `GET /api/persons/:id` - Get a single person
- `POST /api/persons` - Create a new person
- `PUT /api/persons/:id` - Update a person
- `DELETE /api/persons/:id` - Delete a person

## Person Fields

- **Name** (required)
- **Phone** (required)
- **Kids Number** (required, number)
- **Month Income** (required, number)
- **Marital Status** (required: Single, Married, Divorced, Widowed)
- **Live in Renta** (boolean)
- **Has a Car** (boolean)
- **Bank Number** (required)

## Project Structure

```
donation-app/
├── backend/
│   ├── models/
│   │   └── Person.js
│   ├── routes/
│   │   └── persons.js
│   ├── .env
│   ├── package.json
│   └── server.js
├── frontend/
│   ├── public/
│   │   └── index.html
│   ├── src/
│   │   ├── App.js
│   │   └── index.js
│   └── package.json
├── package.json
└── README.md
```

## Technologies Used

- **Frontend**: React Native Web, React
- **Backend**: Node.js, Express
- **Database**: MongoDB with Mongoose
- **Styling**: React Native StyleSheet

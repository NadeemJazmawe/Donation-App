# Database Migration Scripts

This folder contains scripts for managing the MongoDB database for the Donation App.

## Available Scripts

### 1. `migrate.js` - Database Migration
Creates the database and collections if they don't exist, and sets up indexes.

**Usage:**
```bash
cd db-migration
node migrate.js
```

**What it does:**
- Connects to MongoDB
- Creates the database if it doesn't exist
- Creates the `persons` collection
- Sets up indexes on `name`, `phone`, and `createdAt` fields
- Validates the schema

---

### 2. `seed.js` - Seed Sample Data
Populates the database with sample person data for testing.

**Usage:**
```bash
cd db-migration
node seed.js
```

**With force flag (clears existing data first):**
```bash
node seed.js --force
```

**What it does:**
- Inserts 5 sample persons with various data
- Checks if data already exists (won't overwrite unless `--force` is used)
- Displays summary of inserted data

---

### 3. `reset.js` - Reset Database
Clears all data from the database (WARNING: This deletes everything!).

**Usage:**
```bash
cd db-migration
node reset.js
```

**What it does:**
- Deletes all persons from the database
- Shows count of deleted records
- Verifies deletion was successful

---

## Running from Project Root

You can also run these scripts from the project root using npm scripts (if added to package.json):

```bash
# From project root
npm run db:migrate
npm run db:seed
npm run db:reset
```

---

## Prerequisites

1. MongoDB must be running on your local machine
2. Backend `.env` file must be configured (or use default: `mongodb://localhost:27017/donation-app`)
3. Backend dependencies must be installed (`npm install` in backend folder)

---

## Typical Workflow

1. **First time setup:**
   ```bash
   node migrate.js    # Create database and collections
   node seed.js      # Add sample data (optional)
   ```

2. **Reset and reseed:**
   ```bash
   node reset.js     # Clear all data
   node seed.js      # Add fresh sample data
   ```

3. **Force reseed (overwrites existing data):**
   ```bash
   node seed.js --force
   ```

---

## Environment Variables

The scripts use the same MongoDB connection string as the backend:
- Default: `mongodb://localhost:27017/donation-app`
- Can be overridden via `backend/.env` file:
  ```
  MONGODB_URI=mongodb://localhost:27017/donation-app
  ```

---

## Notes

- All scripts connect to MongoDB using the same connection string as the backend
- Scripts automatically close the database connection when finished
- The `persons` collection will be created automatically when the first document is inserted
- Indexes improve query performance on frequently searched fields

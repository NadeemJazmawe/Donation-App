const path = require('path');

// Resolve modules from backend node_modules
const Module = require('module');
const originalRequire = Module.prototype.require;

Module.prototype.require = function(id) {
  // If it's a local module (starts with . or /), use original require
  if (id.startsWith('.') || path.isAbsolute(id)) {
    return originalRequire.apply(this, arguments);
  }
  
  // Try backend node_modules first
  const backendNodeModules = path.join(__dirname, '../backend/node_modules');
  try {
    return originalRequire.call(this, path.join(backendNodeModules, id));
  } catch (err) {
    // Fallback to normal resolution
    return originalRequire.apply(this, arguments);
  }
};

const mongoose = require('mongoose');
const dotenv = require('dotenv');
const Person = require('../backend/models/Person');

dotenv.config({ path: path.join(__dirname, '../backend/.env') });

const MONGODB_URI = process.env.MONGODB_URI || 'mongodb://localhost:27017/donation-app';

/**
 * Database Reset Script
 * This script clears all data from the database
 * WARNING: This will delete all records!
 */
async function reset() {
  try {
    console.log('⚠️  WARNING: This will delete ALL data from the database!');
    console.log('🔄 Starting database reset...');
    console.log(`📡 Connecting to MongoDB: ${MONGODB_URI}`);

    // Connect to MongoDB
    await mongoose.connect(MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    // Get collection count before deletion
    const personsCount = await Person.countDocuments();
    console.log(`📊 Current persons count: ${personsCount}`);

    if (personsCount === 0) {
      console.log('ℹ️  Database is already empty. Nothing to reset.');
      process.exit(0);
    }

    // Delete all documents
    console.log('🗑️  Deleting all persons...');
    const result = await Person.deleteMany({});
    console.log(`✅ Deleted ${result.deletedCount} person(s)`);

    // Verify deletion
    const remainingCount = await Person.countDocuments();
    if (remainingCount === 0) {
      console.log('✅ Database reset completed successfully!');
    } else {
      console.log(`⚠️  Warning: ${remainingCount} person(s) still remain in database.`);
    }

  } catch (error) {
    console.error('❌ Reset failed:', error);
    process.exit(1);
  } finally {
    await mongoose.connection.close();
    console.log('\n🔌 Database connection closed');
  }
}

// Run reset
reset();

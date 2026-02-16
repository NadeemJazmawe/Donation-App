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
 * MongoDB Migration Script
 * This script creates the database and collections if they don't exist
 */
async function migrate() {
  try {
    console.log('🔄 Starting database migration...');
    console.log(`📡 Connecting to MongoDB: ${MONGODB_URI}`);

    // Connect to MongoDB
    await mongoose.connect(MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    // Get the database name from the connection string
    const dbName = mongoose.connection.db.databaseName;
    console.log(`📦 Database: ${dbName}`);

    // Check if database exists
    const adminDb = mongoose.connection.db.admin();
    const databases = await adminDb.listDatabases();
    const dbExists = databases.databases.some(db => db.name === dbName);

    if (dbExists) {
      console.log(`✅ Database '${dbName}' already exists`);
    } else {
      console.log(`📝 Creating database '${dbName}'...`);
      // Database will be created automatically when we create a collection
    }

    // Create collections by inserting and deleting a dummy document
    // This ensures the collection exists with proper indexes
    console.log('📋 Setting up collections...');

    // Create persons collection
    const personsCollection = mongoose.connection.db.collection('persons');
    const personsCount = await personsCollection.countDocuments();
    console.log(`   - 'persons' collection: ${personsCount} documents`);

    // Create indexes for better performance
    console.log('🔍 Creating indexes...');
    await personsCollection.createIndex({ name: 1 });
    await personsCollection.createIndex({ phone: 1 });
    await personsCollection.createIndex({ createdAt: -1 });
    console.log('   ✅ Indexes created on: name, phone, createdAt');

    // Verify the schema by creating a test document and deleting it
    console.log('🧪 Verifying schema...');
    const testPerson = new Person({
      name: '__TEST__',
      phone: '0000000000',
      kidsNumber: 0,
      monthIncome: 0,
      maritalStatus: 'Single',
      liveInRenta: false,
      hasCar: false,
      bankNumber: 'TEST'
    });

    // This will validate the schema
    try {
      await testPerson.validate();
      console.log('   ✅ Schema validation passed');
    } catch (error) {
      console.error('   ❌ Schema validation failed:', error.message);
      throw error;
    }

    console.log('\n✨ Migration completed successfully!');
    console.log(`📊 Database: ${dbName}`);
    console.log(`📋 Collections: persons`);
    console.log(`🔍 Indexes: name, phone, createdAt`);

  } catch (error) {
    console.error('❌ Migration failed:', error);
    process.exit(1);
  } finally {
    await mongoose.connection.close();
    console.log('\n🔌 Database connection closed');
  }
}

// Run migration
migrate();

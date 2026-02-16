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
 * Database Seeding Script
 * This script populates the database with sample data
 */
const samplePersons = [
  {
    name: 'John Doe',
    phone: '1234567890',
    kidsNumber: 2,
    monthIncome: 5000,
    maritalStatus: 'Married',
    liveInRenta: true,
    hasCar: true,
    bankNumber: 'ACC001'
  },
  {
    name: 'Jane Smith',
    phone: '0987654321',
    kidsNumber: 1,
    monthIncome: 3500,
    maritalStatus: 'Single',
    liveInRenta: false,
    hasCar: false,
    bankNumber: 'ACC002'
  },
  {
    name: 'Bob Johnson',
    phone: '5555555555',
    kidsNumber: 3,
    monthIncome: 6000,
    maritalStatus: 'Married',
    liveInRenta: true,
    hasCar: true,
    bankNumber: 'ACC003'
  },
  {
    name: 'Alice Williams',
    phone: '4444444444',
    kidsNumber: 0,
    monthIncome: 4500,
    maritalStatus: 'Divorced',
    liveInRenta: false,
    hasCar: true,
    bankNumber: 'ACC004'
  },
  {
    name: 'Charlie Brown',
    phone: '3333333333',
    kidsNumber: 1,
    monthIncome: 2800,
    maritalStatus: 'Single',
    liveInRenta: true,
    hasCar: false,
    bankNumber: 'ACC005'
  }
];

async function seed() {
  try {
    console.log('🌱 Starting database seeding...');
    console.log(`📡 Connecting to MongoDB: ${MONGODB_URI}`);

    // Connect to MongoDB
    await mongoose.connect(MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    // Check if data already exists
    const existingCount = await Person.countDocuments();
    if (existingCount > 0) {
      console.log(`⚠️  Database already contains ${existingCount} person(s).`);
      console.log('   Use --force flag to clear and reseed, or delete existing data manually.');
      process.exit(0);
    }

    // Insert sample data
    console.log(`📝 Inserting ${samplePersons.length} sample persons...`);
    const inserted = await Person.insertMany(samplePersons);
    console.log(`✅ Successfully inserted ${inserted.length} persons`);

    // Display summary
    console.log('\n📊 Database Summary:');
    console.log(`   Total persons: ${await Person.countDocuments()}`);
    console.log(`   Collections: persons`);

    console.log('\n✨ Seeding completed successfully!');

  } catch (error) {
    console.error('❌ Seeding failed:', error);
    process.exit(1);
  } finally {
    await mongoose.connection.close();
    console.log('\n🔌 Database connection closed');
  }
}

// Check for --force flag
const forceFlag = process.argv.includes('--force');
if (forceFlag) {
  console.log('⚠️  --force flag detected. This will clear all existing data!');
  // Clear existing data
  mongoose.connect(MONGODB_URI).then(async () => {
    await Person.deleteMany({});
    console.log('🗑️  Cleared existing data');
    await mongoose.connection.close();
    // Now run seed
    seed();
  });
} else {
  // Run seed normally
  seed();
}

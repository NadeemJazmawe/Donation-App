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
 * Migration Script: Add status field to existing activities
 * This script updates all existing activities in the database to include a status field
 * Activities without a status will be set to 'pending' by default
 */
async function addActivityStatus() {
  try {
    console.log('🔄 Starting migration: Add status field to activities...');
    console.log(`📡 Connecting to MongoDB: ${MONGODB_URI}`);

    // Connect to MongoDB
    await mongoose.connect(MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    const dbName = mongoose.connection.db.databaseName;
    console.log(`📦 Database: ${dbName}`);

    // Find all persons with activities
    const persons = await Person.find({ activities: { $exists: true, $ne: [] } });
    console.log(`\n📋 Found ${persons.length} person(s) with activities`);

    let totalActivitiesUpdated = 0;
    let totalPersonsUpdated = 0;

    for (const person of persons) {
      let personUpdated = false;
      const updatedActivities = person.activities.map(activity => {
        // If activity doesn't have status field, add it with default 'pending'
        if (!activity.status || (activity.status !== 'pending' && activity.status !== 'completed')) {
          personUpdated = true;
          totalActivitiesUpdated++;
          return {
            ...activity.toObject ? activity.toObject() : activity,
            status: 'pending'
          };
        }
        return activity.toObject ? activity.toObject() : activity;
      });

      if (personUpdated) {
        // Update the person with activities that have status field
        await Person.findByIdAndUpdate(
          person._id,
          { $set: { activities: updatedActivities } },
          { new: true, runValidators: true }
        );
        totalPersonsUpdated++;
        console.log(`   ✅ Updated person: ${person.name} (${updatedActivities.length} activities)`);
      }
    }

    console.log('\n✨ Migration completed successfully!');
    console.log(`📊 Summary:`);
    console.log(`   - Persons processed: ${persons.length}`);
    console.log(`   - Persons updated: ${totalPersonsUpdated}`);
    console.log(`   - Activities updated: ${totalActivitiesUpdated}`);

  } catch (error) {
    console.error('❌ Migration failed:', error);
    process.exit(1);
  } finally {
    await mongoose.connection.close();
    console.log('\n🔌 Database connection closed');
  }
}

// Run migration
addActivityStatus();

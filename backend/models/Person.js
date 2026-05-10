const mongoose = require('mongoose');

// Helper function to calculate age from birthday
function calculateAge(birthday) {
  if (!birthday) return null;
  const today = new Date();
  const birthDate = new Date(birthday);
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

const kidSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    trim: true
  },
  birthday: {
    type: Date,
    required: true
  }
}, { _id: false });

// Virtual for age calculation
kidSchema.virtual('age').get(function() {
  return calculateAge(this.birthday);
});

kidSchema.set('toJSON', { virtuals: true });

const personSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    trim: true
  },
  phone: {
    type: String,
    required: true,
    trim: true
  },
  kidsNumber: {
    type: Number,
    default: 0,
    min: 0
  },
  kids: {
    type: [kidSchema],
    default: []
  },
  monthIncome: {
    type: Number,
    required: false,
    min: 0,
    default: null
  },
  maritalStatus: {
    type: String,
    required: true,
    enum: ['Single', 'Married', 'Divorced', 'Widowed', 'Divorce Project', 'Separated']
  },
  liveInRenta: {
    type: Boolean,
    required: false,
    default: null
  },
  hasCar: {
    type: Boolean,
    required: false,
    default: null
  },
  bankNumber: {
    type: String,
    required: false,
    trim: true,
    default: ''
  },
  description: {
    type: String,
    required: false,
    trim: true,
    default: null
  },
  favorite: {
    type: Boolean,
    default: false
  },
  activities: {
    type: [{
      description: {
        type: String,
        required: true,
        trim: true
      },
      date: {
        type: Date,
        default: Date.now
      },
      status: {
        type: String,
        enum: ['pending', 'completed'],
        default: 'pending'
      },
      distributor: {
        type: String,
        default: '',
        trim: true
      },
      season: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'Season',
        default: null
      }
    }],
    default: []
  }
}, {
  timestamps: true,
  toJSON: { virtuals: true }
});

// Pre-save middleware to auto-calculate kidsNumber from kids array
personSchema.pre('save', function(next) {
  if (this.kids && Array.isArray(this.kids)) {
    this.kidsNumber = this.kids.length;
  } else {
    this.kidsNumber = 0;
  }
  next();
});

// Pre-update middleware for findOneAndUpdate and findByIdAndUpdate
personSchema.pre(['findOneAndUpdate', 'findByIdAndUpdate', 'updateOne'], function(next) {
  const update = this.getUpdate();
  if (update && update.kids !== undefined) {
    if (Array.isArray(update.kids)) {
      update.$set = update.$set || {};
      update.$set.kidsNumber = update.kids.length;
    } else {
      update.$set = update.$set || {};
      update.$set.kidsNumber = 0;
    }
  }
  next();
});

module.exports = mongoose.model('Person', personSchema);

const express = require('express');
const router = express.Router();
const Person = require('../models/Person');

const activityPopulate = { path: 'activities.season', select: 'name' };

// GET all persons
router.get('/', async (req, res) => {
  try {
    const persons = await Person.find().populate(activityPopulate).sort({ createdAt: -1 });
    res.json(persons);
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

// GET single person
router.get('/:id', async (req, res) => {
  try {
    const person = await Person.findById(req.params.id).populate(activityPopulate);
    if (!person) {
      return res.status(404).json({ message: 'Person not found' });
    }
    res.json(person);
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

// POST create new person
router.post('/', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    const phone = (req.body.phone || '').trim();
    if (!name || !phone) {
      return res.status(400).json({ message: 'Name and phone are required' });
    }

    const existing = await Person.findOne({ name, phone });
    if (existing) {
      return res.status(409).json({ message: 'DUPLICATE_PERSON', description: 'A person with this name and phone already exists' });
    }

    // Auto-calculate kidsNumber from kids array if kids are provided
    if (req.body.kids !== undefined) {
      req.body.kidsNumber = Array.isArray(req.body.kids) ? req.body.kids.length : 0;
    } else {
      req.body.kidsNumber = 0;
    }

    if (req.body.activities && Array.isArray(req.body.activities)) {
      req.body.activities = req.body.activities.map(activity => ({
        ...activity,
        status: activity.status || 'pending',
        season: activity.season || null
      }));
    }

    const person = new Person(req.body);
    const savedPerson = await person.save();
    await savedPerson.populate(activityPopulate);
    res.status(201).json(savedPerson);
  } catch (error) {
    res.status(400).json({ message: error.message });
  }
});

// PUT update person
router.put('/:id', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    const phone = (req.body.phone || '').trim();
    if (name && phone) {
      const existing = await Person.findOne({
        name,
        phone,
        _id: { $ne: req.params.id }
      });
      if (existing) {
        return res.status(409).json({ message: 'DUPLICATE_PERSON', description: 'A person with this name and phone already exists' });
      }
    }

    // Auto-calculate kidsNumber from kids array if kids are provided
    if (req.body.kids !== undefined) {
      req.body.kidsNumber = Array.isArray(req.body.kids) ? req.body.kids.length : 0;
    }

    // Ensure activities have status field (default to 'pending' if missing)
    if (req.body.activities && Array.isArray(req.body.activities)) {
      req.body.activities = req.body.activities.map(activity => ({
        ...activity,
        status: activity.status || 'pending',
        season: activity.season || null
      }));
    }

    const person = await Person.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true, runValidators: true }
    ).populate(activityPopulate);
    if (!person) {
      return res.status(404).json({ message: 'Person not found' });
    }
    res.json(person);
  } catch (error) {
    res.status(400).json({ message: error.message });
  }
});

// DELETE person
router.delete('/:id', async (req, res) => {
  try {
    const person = await Person.findByIdAndDelete(req.params.id);
    if (!person) {
      return res.status(404).json({ message: 'Person not found' });
    }
    res.json({ message: 'Person deleted successfully' });
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

module.exports = router;

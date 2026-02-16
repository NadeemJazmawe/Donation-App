const express = require('express');
const router = express.Router();
const Person = require('../models/Person');

// GET all persons
router.get('/', async (req, res) => {
  try {
    const persons = await Person.find().sort({ createdAt: -1 });
    res.json(persons);
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

// GET single person
router.get('/:id', async (req, res) => {
  try {
    const person = await Person.findById(req.params.id);
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
    // Auto-calculate kidsNumber from kids array if kids are provided
    if (req.body.kids !== undefined) {
      req.body.kidsNumber = Array.isArray(req.body.kids) ? req.body.kids.length : 0;
    } else {
      req.body.kidsNumber = 0;
    }
    
    const person = new Person(req.body);
    const savedPerson = await person.save();
    res.status(201).json(savedPerson);
  } catch (error) {
    res.status(400).json({ message: error.message });
  }
});

// PUT update person
router.put('/:id', async (req, res) => {
  try {
    // Auto-calculate kidsNumber from kids array if kids are provided
    if (req.body.kids !== undefined) {
      req.body.kidsNumber = Array.isArray(req.body.kids) ? req.body.kids.length : 0;
    }
    
    // Ensure activities have status field (default to 'pending' if missing)
    if (req.body.activities && Array.isArray(req.body.activities)) {
      req.body.activities = req.body.activities.map(activity => ({
        ...activity,
        status: activity.status || 'pending'
      }));
    }
    
    const person = await Person.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true, runValidators: true }
    );
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

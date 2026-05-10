const express = require('express');
const router = express.Router();
const mongoose = require('mongoose');
const Season = require('../models/Season');
const Person = require('../models/Person');

router.get('/', async (req, res) => {
  try {
    const seasons = await Season.find().sort({ createdAt: -1 });
    res.json(seasons);
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

router.post('/', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    if (!name) {
      return res.status(400).json({ message: 'Season name is required' });
    }
    const season = new Season({ name });
    const saved = await season.save();
    res.status(201).json(saved);
  } catch (error) {
    if (error.code === 11000) {
      return res.status(409).json({ message: 'DUPLICATE_SEASON', description: 'A season with this name already exists' });
    }
    res.status(400).json({ message: error.message });
  }
});

router.delete('/:id', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ message: 'Invalid season id' });
    }
    const inUse = await Person.exists({ 'activities.season': req.params.id });
    if (inUse) {
      return res.status(409).json({
        message: 'SEASON_IN_USE',
        description: 'Cannot delete a season that is assigned to activities'
      });
    }
    const deleted = await Season.findByIdAndDelete(req.params.id);
    if (!deleted) {
      return res.status(404).json({ message: 'Season not found' });
    }
    res.json({ message: 'Season deleted successfully' });
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

module.exports = router;

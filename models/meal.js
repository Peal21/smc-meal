const mongoose = require('mongoose');

const mealSchema = new mongoose.Schema({
  roll: { type: String, required: true },
  batch: { type: String, required: true },
  date: { type: Date, required: true },
  breakfast: { type: Boolean, default: false },
  lunch: { type: Boolean, default: false },
  dinner: { type: Boolean, default: false },
  mutton: { type: Boolean, default: false },
  egg: { type: Boolean, default: false }
});

const Meal = mongoose.model('Meal', mealSchema);

module.exports = Meal;

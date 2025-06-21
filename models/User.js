const mongoose = require('mongoose');

const userSchema = new mongoose.Schema({
  name: { type: String, required: true },
  classRoll: { type: Number, required: true },
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
  gender: { type: String, required: true },
  batch: { type: String, required: true },
  deposit: { type: Number, default: 0 },
  totalMealCount: { type: Number, default: 0 }
});

userSchema.index({ classRoll: 1, batch: 1 }, { unique: true });

module.exports = mongoose.models.User || mongoose.model('User', userSchema);

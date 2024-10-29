const mongoose = require('mongoose');

const userSchema = new mongoose.Schema({
    name: { type: String, required: true },
    email: { type: String, unique: true, required: true },
    password: { type: String, required: true },
    role: { type: String, required: true },
    lunch: { type: String, default: 'Not Selected' },
    dinner: { type: String, default: 'Not Selected' }
});

module.exports = mongoose.model('User', userSchema);

const express = require('express');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const User = require('../models/User');

const router = express.Router();
const secretKey = 'yourSecretKey'; // Change this to a secure key

// Register route
router.post('/register', async (req, res) => {
    const { name, email, password } = req.body;
    try {
        const hashedPassword = await bcrypt.hash(password, 10);
        const newUser = new User({ name, email, password: hashedPassword });
        await newUser.save();
        res.status(201).send('User registered successfully');
    } catch (err) {
        res.status(400).send('Error registering user');
    }
});

// Login route
router.post('/login', async (req, res) => {
    const { email, password } = req.body;
    try {
        const user = await User.findOne({ email });
        if (!user) return res.status(404).send('User not found');

        const isMatch = await bcrypt.compare(password, user.password);
        if (!isMatch) return res.status(400).send('Invalid password');

        const token = jwt.sign({ userId: user._id }, secretKey, { expiresIn: '1h' });
        res.json({ token });
    } catch (err) {
        res.status(400).send('Error logging in');
    }
});

// Update meal selection
router.put('/updateMeal', async (req, res) => {
    const { userId, lunch, dinner } = req.body;
    try {
        await User.findByIdAndUpdate(userId, { lunch, dinner });
        res.send('Meal selection updated');
    } catch (err) {
        res.status(400).send('Error updating meal selection');
    }
});

module.exports = router;

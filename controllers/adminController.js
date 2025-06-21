const User = require('../models/User'); // Mongoose model

exports.loadUsers = async (req, res) => {
  try {
    const users = await User.find({});
    res.json({ success: true, users });
  } catch (error) {
    console.error("Error loading users:", error);
    res.status(500).json({ success: false, message: 'Server Error' });
  }
};

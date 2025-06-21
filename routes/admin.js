const express = require('express');
const router = express.Router();
const bcrypt = require('bcrypt');
const path = require('path');
const Admin = require('../models/Admin');
const User = require('../models/User');
const adminController = require('../controllers/adminController');

// Login page
router.get("/login", (req, res) => {
  res.sendFile(path.join(__dirname, "..", "public", "admin-login.html"));
});

// Login handler
router.post("/login", async (req, res) => {
  const { email, password } = req.body;
  const admin = await Admin.findOne({ email });
  if (!admin || !(await bcrypt.compare(password, admin.password))) {
    return res.status(401).send("Invalid credentials");
  }
  req.session.admin = true;
  res.redirect("/admin/dashboard");
});

// Dashboard page
router.get("/dashboard", (req, res) => {
  if (!req.session.admin) return res.redirect("/admin/login");
  res.sendFile(path.join(__dirname, "..", "public", "admin-dashboard.html"));
});
// Load all users
router.get('/load-users', adminController.loadUsers);

module.exports = router;
// API: get users by gender
router.get("/view/:gender", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Unauthorized");
  const gender = req.params.gender;
  const users = await User.find({ gender }).sort({ batch: 1, classRoll: 1 });
  res.json(users);
});

module.exports = router;

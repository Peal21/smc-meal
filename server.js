const express = require('express');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const bcrypt = require('bcrypt');
const path = require('path');
const cron = require('node-cron');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Verify .env variables
if (!process.env.MONGODB_URI) {
  console.error('Error: MONGODB_URI is not defined in .env file');
  process.exit(1);
}
if (!process.env.SESSION_SECRET) {
  console.warn('Warning: SESSION_SECRET is not defined in .env file. Using default secret.');
}

// MongoDB Connection
const connectDB = async () => {
  try {
    await mongoose.connect(process.env.MONGODB_URI, {
      serverSelectionTimeoutMS: 5000,
    });
    console.log('MongoDB Connected!');
  } catch (err) {
    console.error('MongoDB connection error:', err.message);
    process.exit(1);
  }
};
connectDB();

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(
  session({
    secret: process.env.SESSION_SECRET || 'your-strong-session-secret',
    resave: false,
    saveUninitialized: false,
    store: MongoStore.create({
      mongoUrl: process.env.MONGODB_URI,
      collectionName: 'sessions',
      ttl: 24 * 60 * 60,
    }).on('error', err => console.error('MongoStore error:', err)),
    cookie: {
      secure: process.env.NODE_ENV === 'production' ? true : false,
      maxAge: 24 * 60 * 60 * 1000,
    },
  })
);
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Authentication Middleware
const requireLogin = (req, res, next) => {
  if (!req.session.userId) {
    console.log('Unauthenticated access attempt to', req.originalUrl);
    return res.redirect('/login');
  }
  console.log('Authenticated user:', req.session.userId, 'accessing', req.originalUrl);
  next();
};

const requireAdmin = (req, res, next) => {
  if (!req.session.admin) {
    console.log('Unauthorized admin access attempt');
    return res.redirect('/admin/login');
  }
  next();
};

const requireStaff = (req, res, next) => {
  if (!req.session.staff) {
    console.log('Unauthorized staff access attempt');
    return res.redirect('/staff/login');
  }
  next();
};

// MongoDB Schemas
const userSchema = new mongoose.Schema({
  name: { type: String, required: true },
  classRoll: { type: Number, required: true },
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
  gender: { type: String, enum: ['Male', 'Female'], required: true },
  batch: { type: String, enum: ['09', '10', '11', '12', '13'], required: true },
  deposit: { type: Number, default: 0 },
  totalMealCount: { type: Number, default: 0 },
}, { collection: 'users' });

const mealHistorySchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
  date: { type: Date, required: true },
  meal: { type: String, enum: ['Lunch', 'Dinner', 'Both', 'Off'], required: true },
  additionalItems: [{ type: String }],
  lunchServed: { type: Boolean, default: false },
  dinnerServed: { type: Boolean, default: false },
  dailyMealCount: { type: Number, default: 0 },
}, { collection: 'mealhistories' });

const staffSchema = new mongoose.Schema({
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
}, { collection: 'staff' });

const adminSchema = new mongoose.Schema({
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
}, { collection: 'admins' });

userSchema.index({ classRoll: 1, batch: 1 }, { unique: true });
mealHistorySchema.index({ userId: 1, date: 1 }, { unique: true });

const User = mongoose.model('User', userSchema);
const MealHistory = mongoose.model('MealHistory', mealHistorySchema);
const Staff = mongoose.model('Staff', staffSchema);
const Admin = mongoose.model('Admin', adminSchema);

// Cron Job: Update meal counts and carry forward previous day's meal
cron.schedule('0 0 * * *', async () => {
  try {
    console.log('Running daily meal count update...');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);

    const users = await User.find().lean();
    for (const user of users) {
      let mealHistory = await MealHistory.findOne({ userId: user._id, date: today }).lean();
      if (!mealHistory) {
        const previousMeal = await MealHistory.findOne({ userId: user._id, date: yesterday }).lean();
        if (previousMeal) {
          console.log(`Copying yesterday's meal for user ${user._id}`);
          mealHistory = await new MealHistory({
            userId: user._id,
            date: today,
            meal: previousMeal.meal,
            additionalItems: previousMeal.additionalItems || [],
            dailyMealCount: previousMeal.dailyMealCount,
            lunchServed: false,
            dinnerServed: false,
          }).save();
        } else {
          mealHistory = await new MealHistory({
            userId: user._id,
            date: today,
            meal: 'Off',
            additionalItems: [],
            dailyMealCount: 0,
            lunchServed: false,
            dinnerServed: false,
          }).save();
        }
      }
      const mealCount = mealHistory.meal === 'Lunch' || mealHistory.meal === 'Dinner' ? 1 : mealHistory.meal === 'Both' ? 2 : 0;
      await MealHistory.updateOne(
        { _id: mealHistory._id },
        { dailyMealCount: mealCount }
      );
      await User.updateOne(
        { _id: user._id },
        { $inc: { totalMealCount: mealCount - mealHistory.dailyMealCount } }
      );
    }
    console.log('Daily meal count updated successfully');
  } catch (err) {
    console.error('Error updating daily meal count:', err);
  }
}, {
  scheduled: true,
  timezone: 'Asia/Dhaka'
});

// Routes
app.get('/', (req, res) => {
  console.log('Serving index.html');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/signup', (req, res) => {
  console.log('Serving signup.html');
  res.sendFile(path.join(__dirname, 'public', 'signup.html'));
});

app.get('/login', (req, res) => {
  console.log('Serving login.html');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/admin/login', (req, res) => {
  console.log('Serving admin-login.html');
  res.sendFile(path.join(__dirname, 'public', 'admin-login.html'));
});

app.get('/staff/login', (req, res) => {
  console.log('Serving staff-login.html');
  res.sendFile(path.join(__dirname, 'public', 'staff-login.html'));
});

app.get('/meal-update', requireLogin, async (req, res) => {
  try {
    console.log('Accessing /meal-update, checking session:', req.session.userId);
    const user = await User.findById(req.session.userId).lean();
    if (!user) {
      console.log('User not found:', req.session.userId);
      return res.status(401).send('User not found');
    }
    const filePath = path.join(__dirname, 'public', 'meal-update.html');
    console.log('Attempting to serve meal-update.html from:', filePath);
    res.sendFile(filePath, err => {
      if (err) {
        console.error('Error serving meal-update.html:', err);
        return res.status(500).send('Error loading meal update page');
      }
      console.log('Successfully served meal-update.html for user:', req.session.userId);
    });
  } catch (error) {
    console.error('Error loading meal update page:', error);
    res.status(500).send('Server error');
  }
});

app.get('/meal-update.html', requireLogin, async (req, res) => {
  try {
    console.log('Accessing /meal-update.html, checking session:', req.session.userId);
    const user = await User.findById(req.session.userId).lean();
    if (!user) {
      console.log('User not found:', req.session.userId);
      return res.status(401).send('User not found');
    }
    const filePath = path.join(__dirname, 'public', 'meal-update.html');
    console.log('Attempting to serve meal-update.html from:', filePath);
    res.sendFile(filePath, err => {
      if (err) {
        console.error('Error serving meal-update.html:', err);
        return res.status(500).send('Error loading meal update page');
      }
      console.log('Successfully served meal-update.html for user:', req.session.userId);
    });
  } catch (error) {
    console.error('Error loading meal update page:', error);
    res.status(500).send('Server error');
  }
});

app.get('/meal-dashboard', requireLogin, async (req, res) => {
  try {
    console.log('Accessing /meal-dashboard, checking session:', req.session.userId);
    const user = await User.findById(req.session.userId).lean();
    if (!user) {
      console.log('User not found:', req.session.userId);
      return res.status(401).json({ error: 'User not found' });
    }
    console.log('Serving JSON for user:', req.session.userId);
    res.json({
      name: user.name,
      classRoll: user.classRoll,
      batch: user.batch,
      gender: user.gender,
      totalMealCount: user.totalMealCount,
      deposit: user.deposit,
    });
  } catch (error) {
    console.error('Error loading meal dashboard data:', error);
    res.status(500).json({ error: 'Server error' });
  }
});

app.get('/meal-dashboard.html', requireLogin, async (req, res) => {
  try {
    console.log('Accessing /meal-dashboard.html, checking session:', req.session.userId);
    const user = await User.findById(req.session.userId).lean();
    if (!user) {
      console.log('User not found:', req.session.userId);
      return res.status(401).send('User not found');
    }
    const filePath = path.join(__dirname, 'public', 'meal-dashboard.html');
    console.log('Attempting to serve meal-dashboard.html from:', filePath);
    res.sendFile(filePath, err => {
      if (err) {
        console.error('Error serving meal-dashboard.html:', err);
        return res.status(500).send('Error loading meal dashboard page');
      }
      console.log('Successfully served meal-dashboard.html for user:', req.session.userId);
    });
  } catch (error) {
    console.error('Error loading meal dashboard page:', error);
    res.status(500).send('Server error');
  }
});

app.get('/api/meal-history', requireLogin, async (req, res) => {
  try {
    const mealHistories = await MealHistory.find({ userId: req.session.userId })
      .sort({ date: -1 })
      .lean();
    console.log('Fetched meal history for user:', req.session.userId, 'Count:', mealHistories.length);
    res.json(mealHistories);
  } catch (error) {
    console.error('Error fetching meal history:', error);
    res.status(500).json({ error: 'Failed to fetch meal history' });
  }
});

app.get('/create-admin', async (req, res) => {
  try {
    const existingAdmin = await Admin.findOne({ email: 'admin@example.com' }).lean();
    if (existingAdmin) {
      console.log('Admin already exists:', existingAdmin.email);
      return res.status(400).send('Admin already exists');
    }
    const hashedPassword = await bcrypt.hash('admin123', 10);
    await new Admin({ email: 'admin@example.com', password: hashedPassword }).save();
    console.log('Admin created successfully');
    res.send('Admin created successfully');
  } catch (err) {
    console.error('Error creating admin:', err);
    res.status(500).send('Failed to create admin');
  }
});

app.get('/create-staff', async (req, res) => {
  try {
    console.log('Attempting to create staff account...');
    const existingStaff = await Staff.findOne({ email: 'staff@example.com' }).lean();
    if (existingStaff) {
      console.log('Staff already exists:', existingStaff.email);
      return res.status(400).send('Staff already exists');
    }
    const hashedPassword = await bcrypt.hash('staff123', 10);
    await new Staff({ email: 'staff@example.com', password: hashedPassword }).save();
    console.log('Staff created successfully');
    res.send('Staff created successfully');
  } catch (err) {
    console.error('Error creating staff:', err);
    res.status(500).send('Failed to create staff');
  }
});

app.post('/admin/login', async (req, res) => {
  const { email, password } = req.body;
  console.log('Admin login attempt:', { email });
  try {
    const admin = await Admin.findOne({ email }).lean();
    if (!admin || !(await bcrypt.compare(password, admin.password))) {
      console.log('Invalid admin credentials');
      return res.status(401).send('Invalid credentials');
    }
    req.session.admin = true;
    await req.session.save();
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error('Error during admin login:', err);
    res.status(500).send('Server error');
  }
});

app.post('/staff/login', async (req, res) => {
  const { email, password } = req.body;
  console.log('Staff login attempt:', { email });
  try {
    const staff = await Staff.findOne({ email }).lean();
    if (!staff || !(await bcrypt.compare(password, staff.password))) {
      console.log('Invalid staff credentials');
      return res.status(401).send('Invalid credentials');
    }
    req.session.staff = true;
    await req.session.save();
    res.redirect('/staff/serving');
  } catch (err) {
    console.error('Error during staff login:', err);
    res.status(500).send('Server error');
  }
});

app.post('/signup', async (req, res) => {
  const { name, classRoll, email, password, gender, batch } = req.body;
  console.log('Signup attempt:', { name, classRoll, email, gender, batch });
  try {
    if (!name || !classRoll || !email || !password || !gender || !batch) {
      return res.status(400).send('All fields are required');
    }
    const existingUser = await User.findOne({ email }).lean();
    if (existingUser) {
      console.log('User already exists:', email);
      return res.status(400).send('User already exists. Please log in.');
    }
    if (!['09', '10', '11', '12', '13'].includes(batch)) {
      console.log('Invalid batch:', batch);
      return res.status(400).send('Invalid batch selected.');
    }
    if (!['Male', 'Female'].includes(gender)) {
      console.log('Invalid gender:', gender);
      return res.status(400).send('Invalid gender selected.');
    }
    if (classRoll < 1 || classRoll > 100) {
      console.log('Invalid class roll:', classRoll);
      return res.status(400).send('Class roll must be between 1 and 100.');
    }
    const rollCheck = await User.findOne({ classRoll, batch }).lean();
    if (rollCheck) {
      console.log('Class roll already exists for batch:', { classRoll, batch });
      return res.status(400).send('Class roll already exists for this batch.');
    }
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = new User({ name, classRoll, email, password: hashedPassword, gender, batch });
    await user.save();
    req.session.userId = user._id.toString();
    await req.session.save();
    console.log('User signed up successfully:', user._id);
    res.redirect('/meal-dashboard.html');
  } catch (error) {
    console.error('Error during signup:', error);
    res.status(500).send('Error signing up. Please try again.');
  }
});

app.post('/login', async (req, res) => {
  const { email, password } = req.body;
  console.log('Login attempt:', { email });
  try {
    if (!email || !password) {
      console.log('Missing email or password');
      return res.status(400).send('Email and password are required');
    }
    const user = await User.findOne({ email }).lean();
    if (!user) {
      console.log('User not found:', email);
      return res.status(400).send('Invalid email or password');
    }
    if (!(await bcrypt.compare(password, user.password))) {
      console.log('Invalid password for:', email);
      return res.status(400).send('Invalid email or password');
    }
    req.session.userId = user._id.toString();
    await req.session.save();
    console.log('User logged in successfully:', user._id);
    res.redirect('/meal-dashboard.html');
  } catch (error) {
    console.error('Error during login:', error);
    res.status(500).send('Error logging in. Please try again.');
  }
});

app.post('/meal-update', requireLogin, async (req, res) => {
  let { meal, additionalItems, date } = req.body;
  console.log('Meal update attempt:', { userId: req.session.userId, meal, additionalItems, date });
  try {
    if (!['Lunch', 'Dinner', 'Both', 'Off'].includes(meal)) {
      console.log('Invalid meal type:', meal);
      return res.status(400).json({ error: 'Invalid meal type' });
    }
    if (!date) {
      console.log('Date missing');
      return res.status(400).json({ error: 'Date is required' });
    }
    let additionalItemsArray = [];
    if (Array.isArray(additionalItems)) {
      additionalItemsArray = additionalItems.map(item => item === 'Off' ? 'Egg (Poultry)' : item);
    } else if (typeof additionalItems === 'string' && additionalItems.trim() !== '') {
      additionalItemsArray = [additionalItems === 'Off' ? 'Egg (Poultry)' : additionalItems];
    }
    const selectedDate = new Date(date);
    selectedDate.setHours(0, 0, 0, 0);
    const now = new Date();
    if (new Date(selectedDate.getTime() + 24 * 60 * 60 * 1000 - 1) < now) {
      console.log('Cannot update past date:', selectedDate);
      return res.status(400).json({ error: 'Cannot update meal for past date' });
    }
    const user = await User.findById(req.session.userId).lean();
    if (!user) {
      console.log('User not found:', req.session.userId);
      return res.status(401).json({ error: 'User not found' });
    }
    const existingMeal = await MealHistory.findOne({ userId: req.session.userId, date: selectedDate }).lean();
    const newMealCount = meal === 'Lunch' || meal === 'Dinner' ? 1 : meal === 'Both' ? 2 : 0;
    if (existingMeal) {
      console.log('Updating existing meal:', { meal: existingMeal.meal, dailyMealCount: existingMeal.dailyMealCount });
      const oldMealCount = existingMeal.meal === 'Lunch' || existingMeal.meal === 'Dinner' ? 1 : existingMeal.meal === 'Both' ? 2 : 0;
      await MealHistory.updateOne(
        { _id: existingMeal._id },
        {
          meal,
          additionalItems: additionalItemsArray,
          dailyMealCount: newMealCount,
          lunchServed: meal === 'Both' || meal === 'Lunch' ? existingMeal.lunchServed : false,
          dinnerServed: meal === 'Both' || meal === 'Dinner' ? existingMeal.dinnerServed : false,
        }
      );
      await User.updateOne(
        { _id: req.session.userId },
        { $inc: { totalMealCount: newMealCount - oldMealCount } }
      );
    } else {
      console.log('Creating new meal history for:', { userId: req.session.userId, date: selectedDate });
      await new MealHistory({
        userId: req.session.userId,
        date: selectedDate,
        meal,
        additionalItems: additionalItemsArray,
        dailyMealCount: newMealCount,
        lunchServed: meal === 'Both' || meal === 'Lunch' ? false : false,
        dinnerServed: meal === 'Both' || meal === 'Dinner' ? false : false,
      }).save();
      await User.updateOne(
        { _id: req.session.userId },
        { $inc: { totalMealCount: newMealCount } }
      );
    }
    console.log('Meal updated successfully:', { userId: req.session.userId, meal, dailyMealCount: newMealCount });
    res.json({ message: 'Meal updated successfully' });
  } catch (error) {
    console.error('Error updating meal:', error);
    res.status(500).json({ error: 'Error updating meal. Please try again.' });
  }
});

app.get('/admin/dashboard', requireAdmin, async (req, res) => {
  try {
    const { batch, gender } = req.query;
    let query = {};
    if (batch && batch !== 'all') query.batch = batch;
    if (gender && gender !== 'all') query.gender = gender;
    console.log('Admin dashboard query:', { query });
    const users = await User.find(query).sort({ batch: 1, classRoll: 1 }).lean();
    const batches = ['09', '10', '11', '12', '13'];
    const genders = ['Male', 'Female'];
    res.render('admin-dashboard', {
      users,
      batches,
      genders,
      selectedBatch: batch || 'all',
      selectedGender: gender || 'all',
    });
  } catch (error) {
    console.error('Error loading dashboard:', error);
    res.status(500).send('Server error');
  }
});

app.get('/staff/serving', requireStaff, async (req, res) => {
  try {
    const { batch, gender, date } = req.query;
    let userQuery = {};
    if (batch && batch !== 'all') userQuery.batch = batch;
    if (gender && gender !== 'all') userQuery.gender = gender;
    const users = await User.find(userQuery).sort({ batch: 1, classRoll: 1 }).lean();
    const selectedDate = date ? new Date(date) : new Date();
    selectedDate.setHours(0, 0, 0, 0);
    const mealHistories = await MealHistory.find({ date: selectedDate }).lean();
    const batches = ['09', '10', '11', '12', '13'];
    const genders = ['Male', 'Female'];
    res.render('staff-serving', {
      users,
      mealHistories,
      batches,
      genders,
      selectedBatch: batch || 'all',
      selectedGender: gender || 'all',
      selectedDate: selectedDate.toISOString().split('T')[0],
      isEditable: true, // Always editable for staff
    });
  } catch (error) {
    console.error('Error loading serving page:', error);
    res.status(500).send('Server error');
  }
});

app.post('/api/meal/serve/:userId', requireStaff, async (req, res) => {
  const { userId } = req.params;
  const { mealType, date } = req.body;
  console.log('Serving meal:', { userId, mealType, date });
  try {
    if (!['Lunch', 'Dinner'].includes(mealType)) {
      console.log('Invalid meal type:', mealType);
      return res.status(400).json({ error: 'Invalid meal type' });
    }
    const selectedDate = new Date(date);
    selectedDate.setHours(0, 0, 0, 0);
    const mealHistory = await MealHistory.findOne({ userId, date: selectedDate }).lean();
    if (!mealHistory) {
      console.log('Meal history not found:', { userId, date });
      return res.status(404).json({ error: 'Meal history not found' });
    }
    if (mealHistory.meal === 'Off') {
      console.log('Cannot serve meal for Off status:', { userId, date });
      return res.status(400).json({ error: 'Cannot serve meal for Off status' });
    }
    if (mealType === 'Lunch' && mealHistory.lunchServed) {
      console.log('Lunch already served:', { userId, date });
      return res.status(400).json({ error: 'Lunch already served' });
    }
    if (mealType === 'Dinner' && mealHistory.dinnerServed) {
      console.log('Dinner already served:', { userId, date });
      return res.status(400).json({ error: 'Dinner already served' });
    }
    await MealHistory.updateOne(
      { _id: mealHistory._id },
      { [mealType === 'Lunch' ? 'lunchServed' : 'dinnerServed']: true }
    );
    console.log(`${mealType} served successfully for user:`, userId);
    res.json({ message: `${mealType} served successfully` });
  } catch (err) {
    console.error('Error serving meal:', err);
    res.status(500).json({ error: 'Failed to serve meal' });
  }
});

app.post('/api/meal/extra', requireStaff, async (req, res) => {
  const { date, mealType } = req.body;
  console.log('Adding extra meal:', { date, mealType });
  try {
    if (!['Lunch', 'Dinner', 'Both'].includes(mealType)) {
      console.log('Invalid meal type for extra:', mealType);
      return res.status(400).json({ error: 'Invalid meal type' });
    }
    const selectedDate = new Date(date);
    selectedDate.setHours(0, 0, 0, 0);
    const updated = await MealHistory.updateMany(
      { date: selectedDate, meal: 'Off' },
      { meal: mealType, lunchServed: false, dinnerServed: false, dailyMealCount: mealType === 'Both' ? 2 : 1 }
    );
    if (updated.modifiedCount > 0) {
      const userIds = (await MealHistory.find({ date: selectedDate, meal: mealType }).lean()).map(m => m.userId);
      await User.updateMany(
        { _id: { $in: userIds } },
        { $inc: { totalMealCount: mealType === 'Both' ? 2 : 1 } }
      );
      console.log(`Extra ${mealType} enabled for ${updated.modifiedCount} users`);
    }
    res.json({ message: `Extra ${mealType} enabled for ${updated.modifiedCount} users` });
  } catch (err) {
    console.error('Error enabling extra meals:', err);
    res.status(500).json({ error: 'Failed to enable extra meals' });
  }
});

app.post('/api/meal/staff-update', requireStaff, async (req, res) => {
  const { userId, meal, date } = req.body;
  console.log('Staff meal update attempt:', { userId, meal, date });
  try {
    if (!['Lunch', 'Dinner', 'Both', 'Off'].includes(meal)) {
      console.log('Invalid meal type:', meal);
      return res.status(400).json({ error: 'Invalid meal type' });
    }
    if (!date || !userId) {
      console.log('Missing date or userId');
      return res.status(400).json({ error: 'Date and userId are required' });
    }
    const selectedDate = new Date(date);
    selectedDate.setHours(0, 0, 0, 0);
    const user = await User.findById(userId).lean();
    if (!user) {
      console.log('User not found:', userId);
      return res.status(401).json({ error: 'User not found' });
    }
    const existingMeal = await MealHistory.findOne({ userId, date: selectedDate }).lean();
    const newMealCount = meal === 'Lunch' || meal === 'Dinner' ? 1 : meal === 'Both' ? 2 : 0;
    if (existingMeal) {
      const oldMealCount = existingMeal.meal === 'Lunch' || existingMeal.meal === 'Dinner' ? 1 : existingMeal.meal === 'Both' ? 2 : 0;
      await MealHistory.updateOne(
        { _id: existingMeal._id },
        {
          meal,
          dailyMealCount: newMealCount,
          lunchServed: meal === 'Both' || meal === 'Lunch' ? existingMeal.lunchServed : false,
          dinnerServed: meal === 'Both' || meal === 'Dinner' ? existingMeal.dinnerServed : false,
        }
      );
      await User.updateOne(
        { _id: userId },
        { $inc: { totalMealCount: newMealCount - oldMealCount } }
      );
    } else {
      await new MealHistory({
        userId,
        date: selectedDate,
        meal,
        additionalItems: [],
        dailyMealCount: newMealCount,
        lunchServed: false,
        dinnerServed: false,
      }).save();
      await User.updateOne(
        { _id: userId },
        { $inc: { totalMealCount: newMealCount } }
      );
    }
    console.log('Meal updated by staff successfully:', { userId, meal, dailyMealCount: newMealCount });
    res.json({ message: 'Meal updated successfully' });
  } catch (error) {
    console.error('Error updating meal by staff:', error);
    res.status(500).json({ error: 'Error updating meal. Please try again.' });
  }
});

app.post('/api/users/:id/update', requireAdmin, async (req, res) => {
  const { id } = req.params;
  const { deposit, totalMealCount } = req.body;
  console.log('Updating user:', { id, deposit, totalMealCount });
  try {
    const updates = {};
    if (deposit !== undefined) updates.deposit = Number(deposit);
    if (totalMealCount !== undefined) updates.totalMealCount = Number(totalMealCount);
    if (Object.keys(updates).length === 0) {
      console.log('No updates provided for user:', id);
      return res.status(400).json({ error: 'No updates provided' });
    }
    await User.updateOne({ _id: id }, { $set: updates });
    console.log('User updated successfully:', id);
    res.json({ message: 'User updated successfully' });
  } catch (err) {
    console.error('Error updating user:', err);
    res.status(500).json({ error: 'Failed to update user' });
  }
});

app.get('/api/users', requireAdmin, async (req, res) => {
  try {
    const users = await User.find().sort({ batch: 1, classRoll: 1 }).lean();
    res.json(users);
  } catch (error) {
    console.error('Error fetching users:', error);
    res.status(500).json({ error: 'Failed to fetch users' });
  }
});

app.get('/export-excel', requireLogin, async (req, res) => {
  const loggedInUserId = req.session.userId;
  try {
    const loggedInUser = await User.findById(loggedInUserId).lean();
    if (!loggedInUser) {
      console.log('User not found:', loggedInUserId);
      return res.status(404).send('User not found');
    }
    console.log('Generating Excel for:', { userId: loggedInUserId, batch: loggedInUser.batch, gender: loggedInUser.gender });
    const users = await User.find({ batch: loggedInUser.batch, gender: loggedInUser.gender }).sort({ classRoll: 1 }).lean();
    const mealHistories = await MealHistory.find({ userId: { $in: users.map(u => u._id) } }).lean();
    console.log('Users found:', users.length, 'Meal histories found:', mealHistories.length);
    const workbook = new ExcelJS.Workbook();
    const worksheet = createWorksheet(workbook, loggedInUser.batch, loggedInUser.gender);
    await fillWorksheet(worksheet, users, mealHistories, loggedInUser.batch, loggedInUser.gender);
    const fileName = `Meal_Update_B${loggedInUser.batch}_${loggedInUser.gender}_${new Date().toISOString().split('T')[0]}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error exporting to Excel:', error);
    res.status(500).send('Error exporting to Excel. Please try again.');
  }
});

function createWorksheet(workbook, batch, gender) {
  const worksheet = workbook.addWorksheet(`MUL-B${batch}-${gender.charAt(0)}`);
  worksheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 7 }];
  worksheet.properties.defaultRowHeight = 20;
  worksheet.properties.showGridLines = false;
  return worksheet;
}

async function fillWorksheet(worksheet, users, mealHistories, batch, gender) {
  worksheet.mergeCells('A1:G1');
  worksheet.getCell('A1').value = 'Satkhira Medical College';
  worksheet.getCell('A1').font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getCell('A1').fill = {
    type: 'gradient',
    gradient: 'linear',
    stops: [
      { position: 0, color: { argb: 'FF2E8B57' } },
      { position: 1, color: { argb: 'FF3CB371' } },
    ],
  };
  worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(1).height = 50;

  worksheet.mergeCells('A2:G2');
  worksheet.getCell('A2').value = 'Meal Update List';
  worksheet.getCell('A2').font = { name: 'Arial', size: 14, bold: true };
  worksheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F8FF' } };
  worksheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(2).height = 30;

  worksheet.getCell('A3').value = `Batch: ${batch}`;
  worksheet.getCell('A4').value = `Gender: ${gender}`;
  worksheet.getCell('A5').value = `Date: ${new Date().toLocaleDateString('en-GB')}`;
  worksheet.getCell('A6').value = `Generated: ${new Date().toLocaleString('en-GB', { timeZone: 'Asia/Dhaka' })}`;
  ['A3', 'A4', 'A5', 'A6'].forEach(cell => {
    worksheet.getCell(cell).font = { name: 'Arial', size: 12, bold: true };
    worksheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };
    worksheet.getCell(cell).border = { left: { style: 'thin' }, right: { style: 'thin' } };
  });

  worksheet.getRow(7).values = [
    'Class Roll',
    'Name',
    'Meal',
    'Additional Items',
    'Lunch Served',
    'Dinner Served',
    'Total Meal Count',
  ];
  worksheet.getRow(7).font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getRow(7).fill = {
    type: 'gradient',
    gradient: 'linear',
    stops: [
      { position: 0, color: { argb: 'FF4682B4' } },
      { position: 1, color: { argb: 'FF6495ED' } },
    ],
  };
  worksheet.getRow(7).alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(7).height = 25;
  worksheet.columns = [
    { key: 'classRoll', width: 12 },
    { key: 'name', width: 30 },
    { key: 'meal', width: 15 },
    { key: 'additionalItems', width: 25 },
    { key: 'lunchServed', width: 15 },
    { key: 'dinnerServed', width: 15 },
    { key: 'totalMealCount', width: 18 },
  ];

  worksheet.getRow(7).eachCell(cell => {
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'thin' },
    };
  });

  let rowIndex = 8;
  for (const user of users) {
    const userMeals = mealHistories.filter(mh => mh.userId.toString() === user._id.toString());
    const latestMeal = userMeals.sort((a, b) => new Date(b.date) - new Date(a.date))[0];
    const addItems = latestMeal?.additionalItems?.map(item => item === '' ? '' : item) || [];
    const addItemsStr = addItems.length ? addItems.join(', ') : '-';
    const row = worksheet.addRow({
      classRoll: user.classRoll,
      name: user.name,
      meal: latestMeal?.meal || 'Off',
      additionalItems: addItemsStr,
      lunchServed: latestMeal?.lunchServed ? 'Yes' : 'No',
      dinnerServed: latestMeal?.dinnerServed ? 'Yes' : 'No',
      totalMealCount: user.totalMealCount,
    });
    row.font = { name: 'Arial', size: 10 };
    row.alignment = { vertical: 'middle', horizontal: 'left' };
    row.eachCell(cell => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    const mealCell = row.getCell('meal');
    switch (latestMeal?.meal || 'Off') {
      case 'Lunch':
        mealCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF87CEEB' } };
        break;
      case 'Dinner':
        mealCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDAA520' } };
        break;
      case 'Both':
        mealCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } };
        break;
      case 'Off':
        mealCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCDCDC' } };
        break;
    }
    const addItemsCell = row.getCell('additionalItems');
    if (addItems.includes('')) {
      addItemsCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFB6C1' } };
    } else if (addItems.includes('Egg (Fish)')) {
      addItemsCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFADD8E6' } };
    } else if (addItems.length > 0) {
      addItemsCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFACD' } };
    }
    rowIndex++;
  }

  const totalsRow = worksheet.addRow({ name: 'TOTAL' });
  worksheet.mergeCells(`A${totalsRow.number}:B${totalsRow.number}`);
  totalsRow.font = { name: 'Arial', size: 11, bold: true };
  totalsRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F8FF' } };
  totalsRow.height = 25;
  totalsRow.eachCell(cell => {
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'thin' },
      bottom: 'medium',
      right: 'thin',
    };
  });

  const totals = { Lunch: 0, Dinner: 0, Both: 0, Off: 0, Mutton: 0, EggPoultry: 0, EggFish: 0 };
  let totalDailyMealCount = 0;
  mealHistories.forEach(mh => {
    if (mh.meal === 'Lunch') totals.Lunch++;
    else if (mh.meal === 'Dinner') totals.Dinner++;
    else if (mh.meal === 'Both') totals.Both++;
    else if (mh.meal === 'Off') totals.Off++;
    totalDailyMealCount += mh.dailyMealCount;
    mh.additionalItems?.forEach(item => {
      if (item === 'Mutton') totals.Mutton++;
      else if (item === 'Egg (Poultry)') totals.EggPoultry++;
      else if (item === 'Egg (Fish)') totals.EggFish++;
    });
  });
  const totalRowIndex = rowIndex + 2;
  worksheet.getCell(`A${totalRowIndex}`).value = 'Total Meal Summary:';
  worksheet.getCell(`B${totalRowIndex}`).value = `Lunch: ${totals.Lunch}`;
  worksheet.getCell(`C${totalRowIndex}`).value = `Dinner: ${totals.Dinner}`;
  worksheet.getCell(`D${totalRowIndex}`).value = `Both: ${totals.Both}`;
  worksheet.getCell(`E${totalRowIndex}`).value = `Off: ${totals.Off}`;
  worksheet.getCell(`F${totalRowIndex}`).value = `Mutton: ${totals.Mutton}`;
  worksheet.getCell(`G${totalRowIndex}`).value = `Egg (Poultry): ${totals.EggPoultry}, Egg (Fish): ${totals.EggFish}`;
  const summaryRow = worksheet.getRow(totalRowIndex);
  summaryRow.font = { name: 'Arial', size: 11, bold: true };
  summaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F8FF' } };
  summaryRow.height = 25;
  summaryRow.eachCell(cell => {
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'thin' },
      bottom: 'double',
      right: 'thin',
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  const footerRowIndex = totalRowIndex + 2;
  worksheet.mergeCells(`A${footerRowIndex}:G${footerRowIndex}`);
  worksheet.getCell(`A${footerRowIndex}`).value = 'Generated by Meal Planner System, Satkhira Medical College';
  worksheet.getCell(`A${footerRowIndex}`).font = { name: 'Arial', size: 10, italic: true };
  worksheet.getCell(`A${footerRowIndex}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
  worksheet.getCell(`A${footerRowIndex}`).alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(footerRowIndex).height = 20;
}

// Logout Route
app.get('/logout', (req, res) => {
  req.session.destroy(err => {
    if (err) console.error('Error destroying session:', err);
    res.redirect('/login');
  });
});

// Start server
const server = app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received. Closing server...');
  server.close(() => {
    mongoose.connection.close(false, () => {
      console.log('MongoDB connection closed.');
      process.exit(0);
    });
  });
});
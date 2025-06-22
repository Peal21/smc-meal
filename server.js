const express = require('express');
   const mongoose = require('mongoose');
   const ExcelJS = require('exceljs');
   const session = require('express-session');
   const MongoStore = require('connect-mongo');
   const bcrypt = require('bcrypt');
   const path = require('path');
   const cron = require('node-cron');
   const rateLimit = require('express-rate-limit');
   require('dotenv').config();

   const app = express();
   const PORT = process.env.PORT || 3000;
   const MONGODB_URI = process.env.MONGODB_URI;
   const SESSION_SECRET = process.env.SESSION_SECRET;

   // MongoDB Connection
   async function connectDB() {
     try {
       await mongoose.connect(MONGODB_URI, { serverSelectionTimeoutMS: 30000 });
       console.log('MongoDB Connected at', new Date().toLocaleString('en-US', { timeZone: 'Asia/Dhaka' }));
     } catch (error) {
       console.error('MongoDB connection error:', error.message);
       process.exit(1);
     }
   }
   connectDB();

   // Middleware
   app.use(express.urlencoded({ extended: true }));
   app.use(express.json());
   app.use(express.static(path.join(__dirname, 'public')));
   app.use(rateLimit({
     windowMs: 15 * 60 * 1000, // 15 minutes
     max: 100, // Limit to 100 requests per IP
   }));
   app.use(session({
     secret: SESSION_SECRET,
     resave: false,
     saveUninitialized: false,
     store: MongoStore.create({
       mongoUrl: MONGODB_URI,
       collectionName: 'sessions',
       ttl: 24 * 60 * 60,
     }).on('error', err => console.error('MongoStore error:', err)),
     cookie: {
       secure: process.env.NODE_ENV === 'production',
       maxAge: 24 * 60 * 60 * 1000,
     },
   }));
   app.set('view engine', 'ejs');
   app.set('views', path.join(__dirname, 'views'));

   // Authentication Middleware
   const requireLogin = (req, res, next) => {
     if (!req.session.userId) return res.status(401).json({ error: 'Unauthorized: Please log in' });
     next();
   };

   const requireAdmin = (req, res, next) => {
     if (!req.session.admin) return res.status(401).json({ error: 'Unauthorized: Admin access required' });
     next();
   };

   const requireStaff = (req, res, next) => {
     if (!req.session.staff) return res.status(401).json({ error: 'Unauthorized: Staff access required' });
     next();
   };

   // Schemas
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
     isExtra: { type: Boolean, default: false },
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

   // Cron Job: Update meal counts daily at midnight (Asia/Dhaka)
   cron.schedule('0 0 * * *', async () => {
     try {
       console.log('Running daily meal update at', new Date().toLocaleString('en-US', { timeZone: 'Asia/Dhaka' }));
       const today = new Date();
       today.setHours(0, 0, 0, 0);
       const yesterday = new Date(today);
       yesterday.setDate(today.getDate() - 1);

       const users = await User.find().lean();
       for (const user of users) {
         let mealHistory = await MealHistory.findOne({ userId: user._id, date: today }).lean();
         if (!mealHistory) {
           const previousMeal = await MealHistory.findOne({ userId: user._id, date: yesterday }).lean();
           mealHistory = await new MealHistory({
             userId: user._id,
             date: today,
             meal: previousMeal?.meal || 'Off',
             additionalItems: previousMeal?.additionalItems || [],
             dailyMealCount: previousMeal ? (previousMeal.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(previousMeal.meal) ? 1 : 0) : 0,
             lunchServed: false,
             dinnerServed: false,
           }).save();
         }
         const mealCount = mealHistory.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(mealHistory.meal) ? 1 : 0;
         await MealHistory.updateOne({ _id: mealHistory._id }, { dailyMealCount: mealCount });
         await User.updateOne({ _id: user._id }, { $inc: { totalMealCount: mealCount - mealHistory.dailyMealCount } });
       }
       console.log('Daily meal update completed');
     } catch (error) {
       console.error('Error in cron job:', error.message);
     }
   }, { scheduled: true, timezone: 'Asia/Dhaka' });

   // Routes
   app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

   app.get('/signup', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signup.html')));

   app.get('/login', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));

   app.get('/admin/login', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin-login.html')));

   app.get('/staff/login', (req, res) => res.sendFile(path.join(__dirname, 'public', 'staff-login.html')));

   app.get('/meal-update', requireLogin, (req, res) => res.sendFile(path.join(__dirname, 'public', 'meal-update.html')));

   app.get('/meal-update.html', requireLogin, (req, res) => res.sendFile(path.join(__dirname, 'public', 'meal-update.html')));

   app.get('/meal-dashboard', requireLogin, async (req, res) => {
     try {
       const user = await User.findById(req.session.userId).lean();
       if (!user) return res.status(401).json({ error: 'User not found' });
       res.json({
         name: user.name,
         classRoll: user.classRoll,
         batch: user.batch,
         gender: user.gender,
         totalMealCount: user.totalMealCount,
         deposit: user.deposit,
       });
     } catch (error) {
       console.error('Error fetching dashboard:', error.message);
       res.status(500).json({ error: 'Server error' });
     }
   });

   app.get('/meal-dashboard.html', requireLogin, (req, res) => res.sendFile(path.join(__dirname, 'public', 'meal-dashboard.html')));

   app.get('/api/meal-history', requireLogin, async (req, res) => {
     try {
       const mealHistories = await MealHistory.find({ userId: req.session.userId }).sort({ date: -1 }).lean();
       res.json(mealHistories);
     } catch (error) {
       console.error('Error fetching meal history:', error.message);
       res.status(500).json({ error: 'Failed to fetch meal history' });
     }
   });

   app.get('/create-admin', async (req, res) => {
     try {
       if (await Admin.findOne({ email: 'admin@example.com' }).lean()) return res.status(400).send('Admin already exists');
       const hashedPassword = await bcrypt.hash('admin123', 10);
       await new Admin({ email: 'admin@example.com', password: hashedPassword }).save();
       res.send('Admin created successfully');
     } catch (error) {
       console.error('Error creating admin:', error.message);
       res.status(500).send('Failed to create admin');
     }
   });

   app.get('/create-staff', async (req, res) => {
     try {
       if (await Staff.findOne({ email: 'staff@example.com' }).lean()) return res.status(400).send('Staff already exists');
       const hashedPassword = await bcrypt.hash('staff123', 10);
       await new Staff({ email: 'staff@example.com', password: hashedPassword }).save();
       res.send('Staff created successfully');
     } catch (error) {
       console.error('Error creating staff:', error.message);
       res.status(500).send('Failed to create staff');
     }
   });

   app.post('/admin/login', async (req, res) => {
     const { email, password } = req.body;
     try {
       const admin = await Admin.findOne({ email }).lean();
       if (!admin || !(await bcrypt.compare(password, admin.password))) return res.status(401).send('Invalid credentials');
       req.session.admin = true;
       await req.session.save();
       res.redirect('/admin/dashboard');
     } catch (error) {
       console.error('Error during admin login:', error.message);
       res.status(500).send('Server error');
     }
   });

   app.post('/staff/login', async (req, res) => {
     const { email, password } = req.body;
     try {
       const staff = await Staff.findOne({ email }).lean();
       if (!staff || !(await bcrypt.compare(password, staff.password))) return res.status(401).send('Invalid credentials');
       req.session.staff = true;
       await req.session.save();
       res.redirect('/staff/serving');
     } catch (error) {
       console.error('Error during staff login:', error.message);
       res.status(500).send('Server error');
     }
   });

   app.post('/signup', async (req, res) => {
     const { name, classRoll, email, password, gender, batch } = req.body;
     try {
       if (!name || !classRoll || !email || !password || !gender || !batch) return res.status(400).send('All fields required');
       if (await User.findOne({ email }).lean()) return res.status(400).send('User already exists');
       if (!['09', '10', '11', '12', '13'].includes(batch)) return res.status(400).send('Invalid batch');
       if (!['Male', 'Female'].includes(gender)) return res.status(400).send('Invalid gender');
       if (classRoll < 1 || classRoll > 100) return res.status(400).send('Invalid class roll');
       if (await User.findOne({ classRoll, batch }).lean()) return res.status(400).send('Class roll exists for this batch');
       const hashedPassword = await bcrypt.hash(password, 10);
       const user = await new User({ name, classRoll, email, password: hashedPassword, gender, batch }).save();
       req.session.userId = user._id.toString();
       await req.session.save();
       res.redirect('/meal-dashboard.html');
     } catch (error) {
       console.error('Error during signup:', error.message);
       res.status(500).send('Error signing up');
     }
   });

   app.post('/login', async (req, res) => {
     const { email, password } = req.body;
     try {
       if (!email || !password) return res.status(400).send('Email and password required');
       const user = await User.findOne({ email }).lean();
       if (!user || !(await bcrypt.compare(password, user.password))) return res.status(400).send('Invalid credentials');
       req.session.userId = user._id.toString();
       await req.session.save();
       res.redirect('/meal-dashboard.html');
     } catch (error) {
       console.error('Error during login:', error.message);
       res.status(500).send('Error logging in');
     }
   });

   app.post('/meal-update', requireLogin, async (req, res) => {
     const { meal, additionalItems, date } = req.body;
     try {
       if (!['Lunch', 'Dinner', 'Both', 'Off'].includes(meal)) return res.status(400).json({ error: 'Invalid meal type' });
       if (!date) return res.status(400).json({ error: 'Date required' });
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       if (new Date(selectedDate.getTime() + 24 * 60 * 60 * 1000 - 1) < new Date()) return res.status(400).json({ error: 'Cannot update past date' });
       const user = await User.findById(req.session.userId).lean();
       if (!user) return res.status(401).json({ error: 'User not found' });
       const additionalItemsArray = Array.isArray(additionalItems) ? additionalItems.filter(Boolean) : [additionalItems].filter(Boolean);
       let mealHistory = await MealHistory.findOne({ userId: req.session.userId, date: selectedDate }).lean();
       const newMealCount = meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(meal) ? 1 : 0;
       if (mealHistory) {
         const oldMealCount = mealHistory.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(mealHistory.meal) ? 1 : 0;
         await MealHistory.updateOne({ _id: mealHistory._id }, {
           meal,
           additionalItems: additionalItemsArray,
           dailyMealCount: newMealCount,
           lunchServed: ['Both', 'Lunch'].includes(meal) ? mealHistory.lunchServed : false,
           dinnerServed: ['Both', 'Dinner'].includes(meal) ? mealHistory.dinnerServed : false,
         });
         await User.updateOne({ _id: req.session.userId }, { $inc: { totalMealCount: newMealCount - oldMealCount } });
       } else {
         await new MealHistory({
           userId: req.session.userId,
           date: selectedDate,
           meal,
           additionalItems: additionalItemsArray,
           dailyMealCount: newMealCount,
           lunchServed: false,
           dinnerServed: false,
         }).save();
         await User.updateOne({ _id: req.session.userId }, { $inc: { totalMealCount: newMealCount } });
       }
       res.json({ message: 'Meal updated successfully' });
     } catch (error) {
       console.error('Error updating meal:', error.message);
       res.status(500).json({ error: 'Failed to update meal' });
     }
   });

   app.get('/admin/dashboard', requireAdmin, async (req, res) => {
     const { batch, gender } = req.query;
     try {
       let query = {};
       if (batch && batch !== 'all') query.batch = batch;
       if (gender && gender !== 'all') query.gender = gender;
       const users = await User.find(query).sort({ batch: 1, classRoll: 1 }).lean();
       res.render('admin-dashboard', {
         users,
         batches: ['09', '10', '11', '12', '13'],
         genders: ['Male', 'Female'],
         selectedBatch: batch || 'all',
         selectedGender: gender || 'all',
       });
     } catch (error) {
       console.error('Error loading admin dashboard:', error.message);
       res.status(500).send('Server error');
     }
   });

   app.get('/staff/serving', requireStaff, async (req, res) => {
     const { batch, gender, date } = req.query;
     try {
       let query = {};
       if (batch && batch !== 'all') query.batch = batch;
       if (gender && gender !== 'all') query.gender = gender;
       let users = await User.find(query).sort({ batch: 1, classRoll: 1 }).lean();
       if (!users.length && Object.keys(query).length) users = await User.find().sort({ batch: 1, classRoll: 1 }).lean();
       const selectedDate = date ? new Date(date) : new Date();
       selectedDate.setHours(0, 0, 0, 0);
       let mealHistories = await MealHistory.find({ date: selectedDate }).lean();
       if (mealHistories.length < users.length) {
         for (const user of users) {
           if (!mealHistories.find(mh => mh.userId.toString() === user._id.toString())) {
             const previousMeal = await MealHistory.findOne({ userId: user._id, date: { $lt: selectedDate } }).sort({ date: -1 }).lean();
             await new MealHistory({
               userId: user._id,
               date: selectedDate,
               meal: previousMeal?.meal || 'Off',
               additionalItems: previousMeal?.additionalItems || [],
               dailyMealCount: previousMeal ? (previousMeal.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(previousMeal.meal) ? 1 : 0) : 0,
               lunchServed: false,
               dinnerServed: false,
             }).save();
           }
         }
         mealHistories = await MealHistory.find({ date: selectedDate }).lean();
       }
       res.render('staff-serving', {
         users,
         mealHistories,
         batches: ['09', '10', '11', '12', '13'],
         genders: ['Male', 'Female'],
         selectedBatch: batch || 'all',
         selectedGender: gender || 'all',
         selectedDate: selectedDate.toISOString().split('T')[0],
         isEditable: true,
         error: !users.length ? 'No users found' : null,
       });
     } catch (error) {
       console.error('Error loading staff serving:', error.message);
       res.status(500).render('staff-serving', {
         users: [],
         mealHistories: [],
         batches: ['09', '10', '11', '12', '13'],
         genders: ['Male', 'Female'],
         selectedBatch: batch || 'all',
         selectedGender: gender || 'all',
         selectedDate: new Date().toISOString().split('T')[0],
         isEditable: true,
         error: 'Failed to load data',
       });
     }
   });

   app.post('/api/meal/serve/:userId', requireStaff, async (req, res) => {
     const { userId } = req.params;
     const { mealType, date } = req.body;
     try {
       if (!['Lunch', 'Dinner'].includes(mealType)) return res.status(400).json({ error: 'Invalid meal type' });
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       let mealHistory = await MealHistory.findOne({ userId, date: selectedDate }).lean();
       if (!mealHistory) {
         const user = await User.findById(userId).lean();
         if (!user) return res.status(404).json({ error: 'User not found' });
         const previousMeal = await MealHistory.findOne({ userId, date: { $lt: selectedDate } }).sort({ date: -1 }).lean();
         mealHistory = await new MealHistory({
           userId,
           date: selectedDate,
           meal: previousMeal?.meal || 'Off',
           additionalItems: previousMeal?.additionalItems || [],
           dailyMealCount: previousMeal ? (previousMeal.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(previousMeal.meal) ? 1 : 0) : 0,
           lunchServed: false,
           dinnerServed: false,
         }).save();
       }
       if (mealHistory.meal === 'Off') return res.status(400).json({ error: 'Cannot serve meal for Off status' });
       if ((mealType === 'Lunch' && mealHistory.lunchServed) || (mealType === 'Dinner' && mealHistory.dinnerServed)) return res.status(400).json({ error: `${mealType} already served` });
       if (mealType === 'Lunch' && !['Lunch', 'Both'].includes(mealHistory.meal)) return res.status(400).json({ error: 'Lunch not enabled' });
       if (mealType === 'Dinner' && !['Dinner', 'Both'].includes(mealHistory.meal)) return res.status(400).json({ error: 'Dinner not enabled' });
       await MealHistory.updateOne({ _id: mealHistory._id }, { [mealType === 'Lunch' ? 'lunchServed' : 'dinnerServed']: true });
       res.json({ message: `${mealType} served successfully` });
     } catch (error) {
       console.error('Error serving meal:', error.message);
       res.status(500).json({ error: 'Failed to serve meal' });
     }
   });

   app.post('/api/meal/extra', requireStaff, async (req, res) => {
     const { date, mealType } = req.body;
     try {
       if (!['Lunch', 'Dinner', 'Both'].includes(mealType)) return res.status(400).json({ error: 'Invalid meal type' });
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       const updated = await MealHistory.updateMany({ date: selectedDate, meal: 'Off' }, {
         meal: mealType,
         lunchServed: false,
         dinnerServed: false,
         dailyMealCount: mealType === 'Both' ? 2 : 1,
         isExtra: true,
       });
       if (updated.modifiedCount > 0) {
         const userIds = (await MealHistory.find({ date: selectedDate, meal: mealType }).lean()).map(m => m.userId);
         await User.updateMany({ _id: { $in: userIds } }, { $inc: { totalMealCount: mealType === 'Both' ? 2 : 1 } });
       }
       res.json({ message: `Extra ${mealType} enabled for ${updated.modifiedCount} users` });
     } catch (error) {
       console.error('Error enabling extra meals:', error.message);
       res.status(500).json({ error: 'Failed to enable extra meals' });
     }
   });

   app.post('/api/meal/extra-specific', requireStaff, async (req, res) => {
     const { userId, mealType, date } = req.body;
     try {
       if (!['Lunch', 'Dinner'].includes(mealType)) return res.status(400).json({ error: 'Invalid meal type' });
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       let mealHistory = await MealHistory.findOne({ userId, date: selectedDate }).lean();
       if (!mealHistory) {
         const user = await User.findById(userId).lean();
         if (!user) return res.status(404).json({ error: 'User not found' });
         mealHistory = await new MealHistory({
           userId,
           date: selectedDate,
           meal: 'Off',
           additionalItems: [],
           dailyMealCount: 0,
           lunchServed: false,
           dinnerServed: false,
         }).save();
       }
       let newMeal, newMealCount;
       if (mealHistory.meal === 'Off') {
         newMeal = mealType;
         newMealCount = 1;
       } else if (mealHistory.meal === 'Lunch' && mealType === 'Dinner' || mealHistory.meal === 'Dinner' && mealType === 'Lunch') {
         newMeal = 'Both';
         newMealCount = 2;
       } else {
         return res.status(400).json({ error: `${mealType} already enabled` });
       }
       const oldMealCount = mealHistory.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(mealHistory.meal) ? 1 : 0;
       await MealHistory.updateOne({ _id: mealHistory._id }, {
         meal: newMeal,
         dailyMealCount: newMealCount,
         lunchServed: ['Both', 'Lunch'].includes(newMeal) ? mealHistory.lunchServed : false,
         dinnerServed: ['Both', 'Dinner'].includes(newMeal) ? mealHistory.dinnerServed : false,
         isExtra: true,
       });
       await User.updateOne({ _id: userId }, { $inc: { totalMealCount: newMealCount - oldMealCount } });
       res.json({ message: `Extra ${mealType} enabled` });
     } catch (error) {
       console.error('Error enabling specific extra meal:', error.message);
       res.status(500).json({ error: 'Failed to enable extra meal' });
     }
   });

   app.get('/api/meal/all-users', requireStaff, async (req, res) => {
     const { date, batch, gender } = req.query;
     try {
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       let query = {};
       if (batch && batch !== 'all') query.batch = batch;
       if (gender && gender !== 'all') query.gender = gender;
       const users = await User.find(query).lean();
       const mealHistories = await MealHistory.find({ date: selectedDate }).lean();
       const offUsers = users.map(user => {
         const mealHistory = mealHistories.find(mh => mh.userId.toString() === user._id.toString()) || { meal: 'Off', lunchServed: false, dinnerServed: false };
         const offMeals = [];
         if (mealHistory.meal === 'Off') offMeals.push('Lunch', 'Dinner');
         else if (mealHistory.meal === 'Lunch' && !mealHistory.lunchServed) offMeals.push('Dinner');
         else if (mealHistory.meal === 'Dinner' && !mealHistory.dinnerServed) offMeals.push('Lunch');
         else if (mealHistory.meal === 'Both' && !mealHistory.lunchServed) offMeals.push('Lunch');
         else if (mealHistory.meal === 'Both' && !mealHistory.dinnerServed) offMeals.push('Dinner');
         return offMeals.length ? { _id: user._id, name: user.name, classRoll: user.classRoll, offMeals } : null;
       }).filter(Boolean);
       res.json(offUsers);
     } catch (error) {
       console.error('Error fetching users:', error.message);
       res.status(500).json({ error: 'Failed to fetch users' });
     }
   });

   app.get('/api/meal/total-count', requireStaff, async (req, res) => {
     const { date } = req.query;
     try {
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       const mealHistories = await MealHistory.find({ date: selectedDate }).lean();
       const totalCount = mealHistories.reduce((sum, mh) => sum + mh.dailyMealCount, 0);
       res.json({ totalCount });
     } catch (error) {
       console.error('Error fetching total meal count:', error.message);
       res.status(500).json({ error: 'Failed to fetch total meal count' });
     }
   });

   app.post('/api/meal/staff-update', requireStaff, async (req, res) => {
     const { userId, meal, date } = req.body;
     try {
       if (!['Lunch', 'Dinner', 'Both', 'Off'].includes(meal)) return res.status(400).json({ error: 'Invalid meal type' });
       if (!date || !userId) return res.status(400).json({ error: 'Date and userId required' });
       const selectedDate = new Date(date);
       selectedDate.setHours(0, 0, 0, 0);
       const user = await User.findById(userId).lean();
       if (!user) return res.status(401).json({ error: 'User not found' });
       let mealHistory = await MealHistory.findOne({ userId, date: selectedDate }).lean();
       const newMealCount = meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(meal) ? 1 : 0;
       if (mealHistory) {
         const oldMealCount = mealHistory.meal === 'Both' ? 2 : ['Lunch', 'Dinner'].includes(mealHistory.meal) ? 1 : 0;
         await MealHistory.updateOne({ _id: mealHistory._id }, {
           meal,
           dailyMealCount: newMealCount,
           lunchServed: ['Both', 'Lunch'].includes(meal) ? mealHistory.lunchServed : false,
           dinnerServed: ['Both', 'Dinner'].includes(meal) ? mealHistory.dinnerServed : false,
         });
         await User.updateOne({ _id: userId }, { $inc: { totalMealCount: newMealCount - oldMealCount } });
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
         await User.updateOne({ _id: userId }, { $inc: { totalMealCount: newMealCount } });
       }
       res.json({ message: 'Meal updated successfully' });
     } catch (error) {
       console.error('Error updating meal:', error.message);
       res.status(500).json({ error: 'Failed to update meal' });
     }
   });

   app.post('/api/users/:id/update', requireAdmin, async (req, res) => {
     const { id } = req.params;
     const { deposit, totalMealCount } = req.body;
     try {
       const updates = {};
       if (deposit !== undefined) updates.deposit = Number(deposit);
       if (totalMealCount !== undefined) updates.totalMealCount = Number(totalMealCount);
       if (!Object.keys(updates).length) return res.status(400).json({ error: 'No updates provided' });
       await User.updateOne({ _id: id }, { $set: updates });
       res.json({ message: 'User updated successfully' });
     } catch (error) {
       console.error('Error updating user:', error.message);
       res.status(500).json({ error: 'Failed to update user' });
     }
   });

   app.get('/api/users', requireAdmin, async (req, res) => {
     try {
       const users = await User.find().sort({ batch: 1, classRoll: 1 }).lean();
       res.json(users);
     } catch (error) {
       console.error('Error fetching users:', error.message);
       res.status(500).json({ error: 'Failed to fetch users' });
     }
   });

   app.get('/export-excel', requireLogin, async (req, res) => {
     try {
       const user = await User.findById(req.session.userId).lean();
       if (!user) return res.status(404).send('User not found');
       const users = await User.find({ batch: user.batch, gender: user.gender }).sort({ classRoll: 1 }).lean();
       const mealHistories = await MealHistory.find({ userId: { $in: users.map(u => u._id) } }).lean();
       const workbook = new ExcelJS.Workbook();
       const worksheet = workbook.addWorksheet(`MUL-B${user.batch}-${user.gender.charAt(0)}`);
       worksheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 7 }];
       worksheet.properties.defaultRowHeight = 20;

       worksheet.mergeCells('A1:G1');
       worksheet.getCell('A1').value = 'Satkhira Medical College';
       worksheet.getCell('A1').font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
       worksheet.getCell('A1').fill = { type: 'gradient', gradient: 'linear', stops: [{ position: 0, color: { argb: 'FF2E8B57' } }, { position: 1, color: { argb: 'FF3CB371' } }] };
       worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
       worksheet.getRow(1).height = 50;

       worksheet.mergeCells('A2:G2');
       worksheet.getCell('A2').value = 'Meal Update List';
       worksheet.getCell('A2').font = { name: 'Arial', size: 14, bold: true };
       worksheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };
       worksheet.getRow(2).height = 30;

       worksheet.getCell('A3').value = `Batch: ${user.batch}`;
       worksheet.getCell('A4').value = `Gender: ${user.gender}`;
       worksheet.getCell('A5').value = `Date: ${new Date().toLocaleDateString('en-GB')}`;
       worksheet.getCell('A6').value = `Generated: ${new Date().toLocaleString('en-GB', { timeZone: 'Asia/Dhaka' })}`;
       ['A3', 'A4', 'A5', 'A6'].forEach(cell => {
         worksheet.getCell(cell).font = { name: 'Arial', size: 12, bold: true };
         worksheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };
       });

       worksheet.getRow(7).values = ['Class Roll', 'Name', 'Meal', 'Additional Items', 'Lunch Served', 'Dinner Served', 'Total Meal Count'];
       worksheet.getRow(7).font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
       worksheet.getRow(7).fill = { type: 'gradient', gradient: 'linear', stops: [{ position: 0, color: { argb: 'FF4682B4' } }, { position: 1, color: { argb: 'FF6495ED' } }] };
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

       let rowIndex = 8;
       for (const user of users) {
         const userMeals = mealHistories.filter(mh => mh.userId.toString() === user._id.toString());
         const latestMeal = userMeals.sort((a, b) => new Date(b.date) - new Date(a.date))[0];
         const row = worksheet.addRow({
           classRoll: user.classRoll,
           name: user.name,
           meal: latestMeal?.meal || 'Off',
           additionalItems: latestMeal?.additionalItems?.join(', ') || '-',
           lunchServed: latestMeal?.lunchServed ? 'Yes' : 'No',
           dinnerServed: latestMeal?.dinnerServed ? 'Yes' : 'No',
           totalMealCount: user.totalMealCount,
         });
         row.font = { name: 'Arial', size: 10 };
         row.alignment = { vertical: 'middle', horizontal: 'left' };
         row.eachCell(cell => cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
         rowIndex++;
       }

       const fileName = `Meal_Update_B${user.batch}_${user.gender}_${new Date().toISOString().split('T')[0]}.xlsx`;
       res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
       res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
       await workbook.xlsx.write(res);
       res.end();
     } catch (error) {
       console.error('Error exporting Excel:', error.message);
       res.status(500).send('Error exporting Excel');
     }
   });

   app.get('/logout', (req, res) => {
     req.session.destroy(err => {
       if (err) console.error('Error destroying session:', err.message);
       res.redirect('/login');
     });
   });

   // Start Server
   const server = app.listen(PORT, '0.0.0.0', () => {
     console.log(`Server running on port ${PORT} at ${new Date().toLocaleString('en-US', { timeZone: 'Asia/Dhaka' })}`);
   });

   // Graceful Shutdown
   process.on('SIGTERM', () => {
     console.log('SIGTERM received. Shutting down...');
     server.close(() => {
       mongoose.connection.close(false, () => {
         console.log('MongoDB connection closed');
         process.exit(0);
       });
     });
   });
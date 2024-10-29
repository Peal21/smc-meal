// Required modules
const express = require('express');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const session = require('express-session');
const path = require('path');
const ExcelJS = require('exceljs');
require('dotenv').config(); // Load environment variables

const app = express();
const PORT = process.env.PORT || 3000; // Use Render's PORT or default to 3000

// Connect to MongoDB
async function connectDB() {
    try {
        await mongoose.connect(process.env.MONGODB_URI, {
            useNewUrlParser: true,
            useUnifiedTopology: true
        });
        console.log("MongoDB Connected!");
    } catch (err) {
        console.error("MongoDB connection error:", err);
    }
}

connectDB(); // Call the connection function

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static('public'));
app.use(session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
}));

// User Schema
const userSchema = new mongoose.Schema({
    name: { type: String, required: true },
    classRoll: { type: Number, required: true, unique: true },
    email: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    role: { type: String, required: true },
    totalMealCount: { type: Number, default: 0 }
});

// Meal History Schema
const mealHistorySchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    date: { type: Date, required: true },
    meal: { type: String, required: true }
});

const User = mongoose.model('User', userSchema);
const MealHistory = mongoose.model('MealHistory', mealHistorySchema);

// Routes
// Signup route
app.post('/signup', async (req, res) => {
    const { name, classRoll, email, password, role } = req.body;

    try {
        const existingUser = await User.findOne({ email });
        if (existingUser) return res.status(400).send('User already exists. Please log in.');

        const hashedPassword = await bcrypt.hash(password, 10);
        const user = new User({ name, classRoll, email, password: hashedPassword, role });
        await user.save();

        req.session.userId = user._id;
        req.session.role = user.role;
        res.redirect('/meal-update');
    } catch (error) {
        console.error("Error during signup:", error);
        res.status(500).send('Error signing up. Please try again.');
    }
});

// Login route
app.post('/login', async (req, res) => {
    const { email, password } = req.body;

    try {
        const user = await User.findOne({ email });
        if (!user || !(await bcrypt.compare(password, user.password))) {
            return res.status(400).send('Invalid email or password');
        }

        req.session.userId = user._id;
        req.session.role = user.role;
        res.redirect('/meal-update');
    } catch (error) {
        console.error("Error during login:", error);
        res.status(500).send('Error logging in. Please try again.');
    }
});

// Meal update page route
app.get('/meal-update', (req, res) => {
    if (!req.session.userId) return res.redirect('/login');
    res.sendFile(path.join(__dirname, 'public', 'meal-update.html'));
});

// Handle meal update submissions
app.post('/meal-update', async (req, res) => {
    if (!req.session.userId) {
        return res.status(401).send('You need to log in to update meals');
    }

    const { meal } = req.body;

    try {
        const currentUser = await User.findById(req.session.userId);
        const today = new Date();
        today.setHours(0, 0, 0, 0); // Set time to midnight for today

        // Check if meal history for today exists
        let mealHistoryEntry = await MealHistory.findOne({ userId: currentUser._id, date: today });

        // Determine meal counts based on user input
        if (meal === 'Off') {
            // Reset meal counts for 'Off'
            currentUser.totalMealCount -= mealHistoryEntry ? (mealHistoryEntry.meal === 'Both' ? 2 : 1) : 0; // Adjust total count
            if (mealHistoryEntry) {
                await MealHistory.deleteOne({ _id: mealHistoryEntry._id }); // Remove entry if 'Off'
            }
        } else {
            if (mealHistoryEntry) {
                // Update existing entry for today
                currentUser.totalMealCount -= mealHistoryEntry.meal === 'Both' ? 2 : 1; // Adjust total count
                mealHistoryEntry.meal = meal; // Update meal status
                await mealHistoryEntry.save();
            } else {
                // Create new meal history entry
                mealHistoryEntry = new MealHistory({ userId: currentUser._id, date: today, meal });
                await mealHistoryEntry.save();
            }

            // Increase total meal count based on the selection
            currentUser.totalMealCount += meal === 'Both' ? 2 : 1; // Count meals
        }

        // Save user
        await currentUser.save();
        return res.send('Meal updated successfully');
    } catch (error) {
        console.error("Error updating meal:", error);
        return res.status(500).send('Error updating meal. Please try again.');
    }
});

// Export meal update list to Excel
app.get('/export-excel', async (req, res) => {
    try {
        const users = await User.find().sort({ classRoll: 1 });
        const mealHistory = await MealHistory.find();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Meal Update List');

        worksheet.columns = [
            { header: 'Class Roll', key: 'classRoll', width: 15 },
            { header: 'Name', key: 'name', width: 25 },
            { header: 'Meal Status', key: 'meal', width: 15 },
            { header: 'Daily Count', key: 'dailyCount', width: 15 },
            { header: 'Total Count', key: 'totalCount', width: 15 }
        ];

        let totalDailyMeals = 0;
        let grandTotalMeals = 0;

        users.forEach(user => {
            const today = new Date();
            today.setHours(0, 0, 0, 0); // Set time to midnight for today
            const todayMeal = mealHistory.find(history =>
                history.userId.toString() === user._id.toString() &&
                new Date(history.date).toDateString() === today.toDateString()
            );

            // Calculate daily count and add to totals
            const dailyCount = todayMeal ? (todayMeal.meal === 'Both' ? 2 : 1) : 0;
            totalDailyMeals += dailyCount;
            grandTotalMeals += user.totalMealCount;

            const row = worksheet.addRow({
                classRoll: user.classRoll,
                name: user.name,
                meal: todayMeal ? todayMeal.meal : 'Off',
                dailyCount: dailyCount,
                totalCount: user.totalMealCount
            });

            // Apply color based on meal status
            if (!todayMeal) {
                row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Red for 'Off'
            } else if (todayMeal.meal === 'Both') {
                row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }; // Green for 'Both'
            }
        });

        // Add totals at the bottom
        const totalRow = worksheet.addRow({
            name: 'Total',
            dailyCount: totalDailyMeals,
            totalCount: grandTotalMeals
        });
        totalRow.font = { bold: true };

        // Save the file with today's date
        const todayDate = new Date().toISOString().split('T')[0];
        const fileName = `meal_update_list_${todayDate}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Error exporting to Excel:", error);
        res.status(500).send("Could not export to Excel");
    }
});

// Route to get today's meal status and total meal count for the user
app.get('/get-meal-status', async (req, res) => {
    if (!req.session.userId) return res.status(401).send('You need to log in');

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Set time to midnight for today

    try {
        const user = await User.findById(req.session.userId);
        const todayMeal = await MealHistory.findOne({
            userId: user._id,
            date: today
        });

        const dailyCount = todayMeal ? (todayMeal.meal === 'Both' ? 2 : 1) : 0;

        res.json({
            totalMealCount: user.totalMealCount,
            todayMealStatus: todayMeal ? todayMeal.meal : 'Off',
            dailyCount
        });
    } catch (error) {
        console.error("Error fetching meal status:", error);
        res.status(500).send('Error fetching meal status. Please try again.');
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});

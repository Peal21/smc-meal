// Required modules
const express = require('express');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const session = require('express-session');
const path = require('path');
const ExcelJS = require('exceljs');

// Initialize app and set port
const app = express();
const PORT = process.env.PORT || 3000;

// MongoDB connection URI
const MONGODB_URI = 'mongodb+srv://askpeal121:Peal1234@cluster0.teofx.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0';

// Connect to MongoDB
async function connectDB() {
    try {
        await mongoose.connect(MONGODB_URI, { useNewUrlParser: true, useUnifiedTopology: true });
        console.log("MongoDB Connected!");
    } catch (err) {
        console.error("MongoDB connection error:", err);
        process.exit(1);
    }
}
connectDB(); // Call the connection function

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static('public'));
app.use(session({
    secret: 'your-strong-session-secret',
    resave: false,
    saveUninitialized: false,
}));

// Define User schema and model
const userSchema = new mongoose.Schema({
    name: { type: String, required: true },
    classRoll: { type: Number, required: true },
    email: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    gender: { type: String, required: true },
    batch: { type: String, required: true },
    totalMealCount: { type: Number, default: 0 }
});

// Correctly define a compound index for classRoll and batch
userSchema.index({ classRoll: 1, batch: 1 }, { unique: true });

const User = mongoose.model('User', userSchema);

// Drop any existing single-field index on `classRoll`
User.collection.dropIndex('classRoll_1').catch(err => {
    console.log('No single-field index on classRoll to drop:', err.message);
});

// Define MealHistory schema and model
const mealHistorySchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    date: { type: Date, required: true },
    meal: { type: String, required: true },
    additionalItems: { type: [String], default: [] } // Field for additional items
});
const MealHistory = mongoose.model('MealHistory', mealHistorySchema); // Define MealHistory model

// Routes
// Signup route
app.post('/signup', async (req, res) => {
    const { name, classRoll, email, password, gender, batch } = req.body;

    console.log("Signup data received:", req.body); // For debugging

    try {
        // Check if the email already exists
        const existingUserByEmail = await User.findOne({ email });
        if (existingUserByEmail) {
            return res.status(400).send('User already exists. Please log in.');
        }

        // Validate the batch
        if (!['09', '10', '11', '12', '13'].includes(batch)) {
            return res.status(400).send('Invalid batch selected.');
        }

        // Validate classRoll range
        if (classRoll < 1 || classRoll > 100) {
            return res.status(400).send('Class roll must be between 1 and 100.');
        }

        // Check if the classRoll already exists in the same batch
        const existingUserByClassRoll = await User.findOne({ classRoll, batch });
        if (existingUserByClassRoll) {
            console.log(`Duplicate roll found for batch ${batch}:`, classRoll); // For debugging
            return res.status(400).send('Class roll already exists. Please choose a different one.');
        }

        // Hash the password
        const hashedPassword = await bcrypt.hash(password, 10);

        // Create a new user
        const user = new User({ name, classRoll, email, password: hashedPassword, gender, batch });
        await user.save();

        // Store user ID in session
        req.session.userId = user._id;

        // Redirect to meal update page
        res.redirect('/meal-update');
    } catch (error) {
        console.error("Error during signup:", error); // For debugging
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
    if (!req.session.userId) return res.status(401).send('You need to log in to update meals');

    const { meal, additionalItems } = req.body; // Capture additionalItems
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset time to midnight for date comparison

    try {
        const currentUser = await User.findById(req.session.userId);
        let mealHistoryEntry = await MealHistory.findOne({ userId: currentUser._id, date: today });

        if (mealHistoryEntry) {
            // If an entry already exists for today, update it
            currentUser.totalMealCount -= (mealHistoryEntry.meal === 'Both' ? 2 : 1);
            mealHistoryEntry.meal = meal;
            mealHistoryEntry.additionalItems = additionalItems; // Save additional items
            await mealHistoryEntry.save();
            currentUser.totalMealCount += (meal === 'Both' ? 2 : 1);
        } else {
            // No entry for today, create a new one
            mealHistoryEntry = new MealHistory({
                userId: currentUser._id,
                date: today,
                meal,
                additionalItems
            });
            await mealHistoryEntry.save();
            currentUser.totalMealCount += (meal === 'Both' ? 2 : 1);
        }

        await currentUser.save();
        res.send('Meal updated successfully');
    } catch (error) {
        console.error("Error updating meal:", error);
        res.status(500).send('Error updating meal. Please try again.');
    }
});

// Export to Excel route
app.get('/export-excel', async (req, res) => {
    try {
        const loggedInUserId = req.session.userId;
        if (!loggedInUserId) {
            return res.status(401).send('User not authenticated');
        }

        const loggedInUser = await User.findById(loggedInUserId);
        const users = await User.find({
            batch: loggedInUser.batch,
            gender: loggedInUser.gender
        }).sort({ classRoll: 1 });

        const mealHistory = await MealHistory.find();

        // Create a workbook for this user's batch and gender only
        const workbook = new ExcelJS.Workbook();
        const worksheet = createWorksheet(workbook, loggedInUser.batch, loggedInUser.gender);

        // Fill the worksheet with user data for this batch and gender
        await fillWorksheet(worksheet, users, mealHistory);

        // Set the response headers to prompt download
        const fileName = `Meal_Update_List_B${loggedInUser.batch}_${loggedInUser.gender}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Write the workbook to the response directly
        await workbook.xlsx.write(res);
        res.end(); // End the response to signal download completion

    } catch (error) {
        console.error("Error exporting to Excel:", error);
        res.status(500).send('Error exporting to Excel. Please try again.');
    }
});

// Function to create a worksheet with specified headers
function createWorksheet(workbook, batch, gender) {
    const worksheet = workbook.addWorksheet(`MUL-B${batch}-${gender.charAt(0)}`);
    worksheet.columns = [
        { header: 'Class Roll', key: 'classRoll', width: 15 },
        { header: 'Name', key: 'name', width: 25 },
        { header: 'Meal Status', key: 'meal', width: 15 },
        { header: 'Mutton', key: 'mutton', width: 15 },
        { header: 'Egg (instead of Fish)', key: 'egg', width: 15 },
        { header: 'Egg instead of poultry', key: 'off', width: 15 },
        { header: 'Daily Count', key: 'dailyCount', width: 15 }
    ];
    return worksheet;
}

// Function to fill the worksheet with user data
async function fillWorksheet(worksheet, users, mealHistory) {
    let totalMutton = 0;
    let totalEgg = 0;
    let totalOff = 0;
    let totalDailyMeals = 0;
    let totalMealCount = 0;

    for (const user of users) {
        const userMeals = mealHistory.filter(entry => entry.userId.equals(user._id));
        const dailyCounts = userMeals.map(meal => ({
            mutton: meal.additionalItems.includes('Mutton') ? 1 : 0,
            egg: meal.additionalItems.includes('Egg') ? 1 : 0,
            off: meal.additionalItems.includes('Off') ? 1 : 0
        }));

        const dailyCount = dailyCounts.reduce((acc, curr) => ({
            mutton: acc.mutton + curr.mutton,
            egg: acc.egg + curr.egg,
            off: acc.off + curr.off,
        }), { mutton: 0, egg: 0, off: 0 });

        // Display meal status and apply color formatting
        let mealStatus = 'Off';
        if (user.totalMealCount === 2) {
            mealStatus = 'Both';
        } else if (user.totalMealCount === 1) {
            mealStatus = user.mealType === 'Lunch' ? 'Lunch' : 'Dinner';
        }

        const row = worksheet.addRow({
            classRoll: user.classRoll,
            name: user.name,
            meal: mealStatus,
            mutton: dailyCount.mutton,
            egg: dailyCount.egg,
            off: dailyCount.off,
            dailyCount: user.totalMealCount
        });

        // Apply specific colors for meal status
        if (mealStatus === 'Both') {
            row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '00FF00' } }; // Green for 'Both'
        } else if (mealStatus === 'Lunch') {
            row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } }; // Yellow for 'Lunch'
        } else if (mealStatus === 'Dinner') {
            row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA07A' } }; // Light Salmon for 'Dinner'
        } else if (mealStatus === 'Off') {
            row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '808080' } }; // Grey for 'Off'
        }

        // Color coding for additional items
        if (dailyCount.mutton > 0) {
            row.getCell('mutton').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0CB' } }; // Pink for Mutton
        }
        if (dailyCount.egg > 0) {
            row.getCell('egg').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF99' } }; // Light Yellow for Egg
        }
        if (dailyCount.off > 0) {
            row.getCell('off').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EE82EE' } }; // Violet for Off
        }

        // Sum totals
        totalDailyMeals += dailyCount.mutton + dailyCount.egg + dailyCount.off;
        totalMutton += dailyCount.mutton;
        totalEgg += dailyCount.egg;
        totalOff += dailyCount.off;
        totalMealCount += user.totalMealCount;
    }

    // Add totals at the end of the worksheet
    worksheet.addRow({
        classRoll: 'Total',
        name: '',
        meal: '',
        mutton: totalMutton,
        egg: totalEgg,
        off: totalOff,
        dailyCount: totalDailyMeals
    });
}


// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

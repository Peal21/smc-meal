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
        await mongoose.connect(MONGODB_URI);
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
    classRoll: { type: Number, required: true, unique: true },
    email: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    role: { type: String, required: true },
    totalMealCount: { type: Number, default: 0 }
});
const User = mongoose.model('User', userSchema);

// Define MealHistory schema and model
const mealHistorySchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    date: { type: Date, required: true },
    meal: { type: String, required: true },
    additionalItems: { type: [String], default: [] } // Field for additional items
});
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
            { header: 'Mutton', key: 'mutton', width: 15 },
            { header: 'Egg insted of fish', key: 'egg', width: 15 },
            { header: 'Egg instead of poultry', key: 'off', width: 15 },
            { header: 'Daily Count', key: 'dailyCount', width: 15 },
            { header: 'Total Count', key: 'totalCount', width: 15 }
        ];

        let totalDailyMeals = 0;
        let grandTotalMeals = 0;
        let totalMutton = 0;
        let totalEgg = 0;
        let totalOff = 0;

        users.forEach(user => {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const todayMeal = mealHistory.find(history =>
                history.userId.toString() === user._id.toString() &&
                new Date(history.date).toDateString() === today.toDateString()
            );

            const dailyCount = todayMeal ? (todayMeal.meal === 'Both' ? 2 : 1) : 0;
            totalDailyMeals += dailyCount;
            grandTotalMeals += user.totalMealCount;

            const row = worksheet.addRow({
                classRoll: user.classRoll,
                name: user.name,
                meal: todayMeal ? todayMeal.meal : 'Off',
                mutton: todayMeal && todayMeal.additionalItems.includes('Mutton') ? 1 : 0,
                egg: todayMeal && todayMeal.additionalItems.includes('Egg') ? 1 : 0,
                off: todayMeal && todayMeal.additionalItems.includes('Off') ? 1 : 0,
                dailyCount: dailyCount,
                totalCount: user.totalMealCount
            });

            // Color meal options
            if (todayMeal) {
                if (todayMeal.meal === 'Lunch') {
                    row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFADD8E6' } }; // Light Blue
                } else if (todayMeal.meal === 'Dinner') {
                    row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } }; // Orange
                } else if (todayMeal.meal === 'Both') {
                    row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }; // Green
                }
            } else {
                row.getCell('meal').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Red for Off
            }

            // Color additional items
            if (todayMeal) {
                if (todayMeal.additionalItems.includes('Mutton')) {
                    row.getCell('mutton').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC0CB' } }; // Pink
                    totalMutton++;
                }
                if (todayMeal.additionalItems.includes('Egg')) {
                    row.getCell('egg').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
                    totalEgg++;
                }
                if (todayMeal.additionalItems.includes('Off')) {
                    row.getCell('off').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEE82EE' } }; // Violet
                    totalOff++;
                }
            }
        });

        // Total summary row
        const totalRow = worksheet.addRow({
            classRoll: 'Total',
            name: '',
            meal: '',
            mutton: totalMutton,
            egg: totalEgg,
            off: totalOff,
            dailyCount: totalDailyMeals,
            totalCount: grandTotalMeals
        });

        // Style total summary row
        if (totalRow) {
            totalRow.eachCell(cell => {
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } }; // Gray for summary
            });
        }

        // Set headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="meal_update_list.xlsx"');

        // Write workbook to response
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Error exporting to Excel:", error);
        res.status(500).send('Error exporting to Excel');
    }
});



// Logout route
app.post('/logout', (req, res) => {
    req.session.destroy(err => {
        if (err) return res.status(500).send('Error logging out');
        res.redirect('/index.html'); // Redirect to the main index page after logout
    });
});
// Start server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

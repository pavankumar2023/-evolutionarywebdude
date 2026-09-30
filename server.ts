import express from 'express';
import path from 'path';
import fs from 'fs';
import { createServer as createViteServer } from 'vite';
import { ExcelStorageEngine } from './server/excelStorage.js';

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ extended: true, limit: '50mb' }));

  // Initialize Excel Storage Engine
  const excelDb = ExcelStorageEngine.getInstance();

  // --- API ROUTES ---

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({
      status: 'ok',
      service: 'Evolutionary Web Dude Backend Service',
      storage: 'Excel File Database (ewd_data_store.xlsx)',
      timestamp: new Date().toISOString()
    });
  });

  // 1. AUTHENTICATION (Backed by Excel Users Worksheet)
  app.post('/api/auth/register', async (req, res) => {
    try {
      const { name, email, password, phone, organization, role } = req.body;
      if (!email || !password) {
        return res.status(400).json({ error: 'Email and password are required' });
      }

      const existing = await excelDb.getUserByEmail(email);
      if (existing) {
        return res.status(409).json({ error: 'User already exists with this email address' });
      }

      const newUser = await excelDb.createUser({
        name: name || email.split('@')[0],
        email,
        password,
        phone: phone || '',
        organization: organization || '',
        role: role === 'ADMIN' ? 'ADMIN' : 'USER'
      });

      const { password: _, ...safeUser } = newUser;
      res.status(201).json({
        message: 'Account registered successfully in Excel data store',
        user: safeUser,
        token: `ewd-jwt-token-${safeUser.id}-${Date.now()}`
      });
    } catch (err: any) {
      console.error('Register error:', err);
      res.status(500).json({ error: err.message || 'Internal server error' });
    }
  });

  app.post('/api/auth/login', async (req, res) => {
    try {
      const { email, password, role } = req.body;
      if (!email || !password) {
        return res.status(400).json({ error: 'Email and password are required' });
      }

      const user = await excelDb.getUserByEmail(email);
      if (!user) {
        return res.status(401).json({ error: 'Invalid email or password' });
      }

      if (user.password !== password) {
        return res.status(401).json({ error: 'Invalid email or password' });
      }

      if (role === 'ADMIN' && user.role !== 'ADMIN') {
        return res.status(403).json({ error: 'Access denied: Admin credentials required' });
      }

      const { password: _, ...safeUser } = user;
      res.json({
        message: 'Authentication successful from Excel store',
        user: safeUser,
        token: `ewd-jwt-token-${safeUser.id}-${Date.now()}`
      });
    } catch (err: any) {
      console.error('Login error:', err);
      res.status(500).json({ error: err.message || 'Internal server error' });
    }
  });

  app.get('/api/auth/users', async (req, res) => {
    try {
      const users = await excelDb.getAllUsers();
      const safeUsers = users.map(({ password, ...u }) => u);
      res.json(safeUsers);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/auth/profile/update', async (req, res) => {
    try {
      const { id, name, phone, organization, bio, avatar, password, role } = req.body;
      if (!id) return res.status(400).json({ error: 'User ID is required' });

      const updated = await excelDb.updateUser(id, { name, phone, organization, bio, avatar, password, role });
      if (!updated) return res.status(404).json({ error: 'User not found in Excel' });

      const { password: _, ...safeUser } = updated;
      res.json({ message: 'Profile updated in Excel sheet', user: safeUser });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/auth/users/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteUser(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'User record not found' });
      res.json({ message: 'User removed from Excel Users store' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 2. COURSES / TRAINING PROGRAMS
  app.get('/api/courses', async (req, res) => {
    try {
      const courses = await excelDb.getAllCourses();
      res.json(courses);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/courses', async (req, res) => {
    try {
      const created = await excelDb.createCourse(req.body);
      res.status(201).json({ message: 'Course created in Excel store', course: created });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/courses/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateCourse(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Course not found' });
      res.json({ message: 'Course updated in Excel store', course: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/courses/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteCourse(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Course not found' });
      res.json({ message: 'Course deleted from Excel store' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 3. ENROLLMENTS
  app.get('/api/enrollments', async (req, res) => {
    try {
      const userEmail = req.query.email as string;
      const all = await excelDb.getAllEnrollments();
      if (userEmail) {
        return res.json(all.filter(e => e.userEmail.toLowerCase() === userEmail.toLowerCase()));
      }
      res.json(all);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/enrollments', async (req, res) => {
    try {
      const newEnrollment = await excelDb.createEnrollment(req.body);
      res.status(201).json({
        message: 'Enrollment registered and stored in Excel sheet',
        enrollment: newEnrollment
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/enrollments/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateEnrollment(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Enrollment not found' });
      res.json({ message: 'Enrollment updated in Excel', enrollment: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/enrollments/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteEnrollment(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Enrollment record not found' });
      res.json({ message: 'Enrollment record removed from Excel' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 4. CONTACTS / LEADS
  app.get('/api/contacts', async (req, res) => {
    try {
      const contacts = await excelDb.getAllContacts();
      res.json(contacts);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/contacts', async (req, res) => {
    try {
      const newContact = await excelDb.createContact(req.body);
      res.status(201).json({
        message: 'Your inquiry has been submitted and recorded in our Excel CRM pipeline',
        contact: newContact
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/contacts/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateContact(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Contact not found' });
      res.json({ message: 'Contact status updated in Excel', contact: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/contacts/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteContact(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Contact record not found' });
      res.json({ message: 'Contact record removed from Excel' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 5. APPOINTMENTS / CONSULTATIONS
  app.get('/api/appointments', async (req, res) => {
    try {
      const appts = await excelDb.getAllAppointments();
      res.json(appts);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/appointments', async (req, res) => {
    try {
      const newAppt = await excelDb.createAppointment(req.body);
      res.status(201).json({
        message: 'Consultation appointment scheduled and recorded in Excel store',
        appointment: newAppt
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/appointments/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateAppointment(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Appointment not found' });
      res.json({ message: 'Appointment updated in Excel', appointment: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/appointments/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteAppointment(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Appointment not found' });
      res.json({ message: 'Appointment removed from Excel' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 6. SHOP PRODUCTS & ORDERS
  app.get('/api/products', (req, res) => {
    const products = [
      {
        id: 'prod-paperless-suite',
        title: 'EWD Paperless Office Suite',
        category: 'Enterprise Solution',
        description: 'Complete digital portal software for automating paper forms, academic student thesis records, and approval workflows.',
        price: 499,
        originalPrice: 799,
        badge: 'Bestseller',
        rating: 4.9,
        reviewsCount: 38,
        features: ['Automated PDF & Digital Signatures', 'Multi-level Role Approvals', 'Real-time Excel & Analytics Sync', 'Self-Optimizing UI Engine'],
        techStack: ['React 18', 'Spring Boot 3', 'Excel Data Pipeline', 'JWT Security']
      },
      {
        id: 'prod-adaptive-cms',
        title: 'Adaptive Web Platform Starter Kit',
        category: 'SaaS Template',
        description: 'Next-generation web portal boilerplate featuring evolutionary layout optimization algorithms and modern animated UI.',
        price: 249,
        originalPrice: 399,
        badge: 'Popular',
        rating: 4.8,
        reviewsCount: 52,
        features: ['Framer Motion Animations', 'Responsive Tailwind CSS 4 Design', 'REST API Architecture', 'Pre-configured Auth & Roles'],
        techStack: ['React', 'TypeScript', 'Tailwind', 'Motion']
      },
      {
        id: 'prod-ai-analytics-pack',
        title: 'Intelligent AI Telemetry & Lead Engine',
        category: 'Starter Kit',
        description: 'Embedded client micro-interaction tracking and automated lead qualification assistant for enterprise portals.',
        price: 349,
        originalPrice: 500,
        rating: 5.0,
        reviewsCount: 24,
        features: ['User Interaction Telemetry', 'Automated Lead Scoring', 'Google Meet / Calendar Booking Sync', 'Excel Export Engine'],
        techStack: ['Node.js / Python', 'AI Analytics', 'ExcelJS']
      },
      {
        id: 'prod-security-suite',
        title: 'Zero-Trust Enterprise Security Pack',
        category: 'Security Suite',
        description: 'Hardened cybersecurity middleware featuring encrypted tokens, brute-force throttling, and role-based audit trail.',
        price: 399,
        originalPrice: 599,
        badge: 'Enterprise',
        rating: 4.9,
        reviewsCount: 19,
        features: ['JWT Refresh Rotation', 'Role-Based Access Control', 'Encrypted Excel Export', 'Security Audit Logs'],
        techStack: ['Spring Security / Express', 'Crypto', 'RBAC']
      }
    ];
    res.json(products);
  });

  app.get('/api/orders', async (req, res) => {
    try {
      const userEmail = req.query.email as string;
      const allOrders = await excelDb.getAllOrders();
      if (userEmail) {
        return res.json(allOrders.filter(o => o.customerEmail.toLowerCase() === userEmail.toLowerCase()));
      }
      res.json(allOrders);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/orders', async (req, res) => {
    try {
      const order = await excelDb.createOrder(req.body);
      res.status(201).json({
        message: 'Order successfully processed and stored in Excel Orders worksheet',
        order
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 7. TESTIMONIALS
  app.get('/api/testimonials', async (req, res) => {
    try {
      const all = await excelDb.getAllTestimonials();
      res.json(all);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/testimonials', async (req, res) => {
    try {
      const newTestimonial = await excelDb.createTestimonial(req.body);
      res.status(201).json({
        message: 'Thank you! Your testimonial has been stored in our Excel records',
        testimonial: newTestimonial
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/testimonials/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateTestimonial(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Testimonial not found' });
      res.json({ message: 'Testimonial updated in Excel', testimonial: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/testimonials/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteTestimonial(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Testimonial not found' });
      res.json({ message: 'Testimonial deleted from Excel' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 8. NEWSLETTER
  app.get('/api/newsletter', async (req, res) => {
    try {
      const subs = await excelDb.getAllSubscribers();
      res.json(subs);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/newsletter/subscribe', async (req, res) => {
    try {
      const { email } = req.body;
      if (!email) return res.status(400).json({ error: 'Email is required' });
      const sub = await excelDb.subscribeNewsletter(email);
      res.status(201).json({
        message: 'Subscribed successfully! Recorded in Excel Newsletter worksheet',
        subscriber: sub
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 9. BLOGS (Add, Edit, Delete, List)
  app.get('/api/blogs', async (req, res) => {
    try {
      const blogs = await excelDb.getAllBlogs();
      res.json(blogs);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/blogs', async (req, res) => {
    try {
      const newBlog = await excelDb.createBlog(req.body);
      res.status(201).json({
        message: 'Blog article successfully published and saved in Excel store',
        blog: newBlog
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/blogs/:id', async (req, res) => {
    try {
      const updated = await excelDb.updateBlog(req.params.id, req.body);
      if (!updated) return res.status(404).json({ error: 'Blog article not found' });
      res.json({ message: 'Blog article updated in Excel store', blog: updated });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/blogs/:id', async (req, res) => {
    try {
      const deleted = await excelDb.deleteBlog(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Blog article not found' });
      res.json({ message: 'Blog article deleted from Excel store' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 10. LIVE EXCEL FILE DATA MANAGEMENT & DOWNLOAD
  app.get('/api/excel/download', (req, res) => {
    const filePath = excelDb.getFilePath();
    if (fs.existsSync(filePath)) {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="EvolutionaryWebDude_Database.xlsx"');
      const fileStream = fs.createReadStream(filePath);
      fileStream.pipe(res);
    } else {
      res.status(404).json({ error: 'Excel file not found on server' });
    }
  });

  app.get('/api/excel/sheets', async (req, res) => {
    try {
      const sheets = await excelDb.getRawSheetsOverview();
      res.json({
        filePath: 'ewd_data_store.xlsx',
        fileSize: fs.existsSync(excelDb.getFilePath()) ? `${Math.round(fs.statSync(excelDb.getFilePath()).size / 1024)} KB` : 'N/A',
        totalSheets: sheets.length,
        sheets
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/excel/cell/update', async (req, res) => {
    try {
      const { sheetName, rowNumber, columnKey, newValue } = req.body;
      if (!sheetName || !rowNumber || !columnKey) {
        return res.status(400).json({ error: 'sheetName, rowNumber, and columnKey are required' });
      }

      const success = await excelDb.updateCellInSheet(sheetName, Number(rowNumber), columnKey, newValue);
      if (!success) {
        return res.status(400).json({ error: 'Failed to update Excel cell' });
      }

      res.json({ message: 'Cell updated directly in Excel file' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // 11. ADMIN DASHBOARD AGGREGATED METRICS
  app.get('/api/admin/stats', async (req, res) => {
    try {
      const [users, courses, enrollments, contacts, appointments, orders, testimonials, subscribers, blogs] = await Promise.all([
        excelDb.getAllUsers(),
        excelDb.getAllCourses(),
        excelDb.getAllEnrollments(),
        excelDb.getAllContacts(),
        excelDb.getAllAppointments(),
        excelDb.getAllOrders(),
        excelDb.getAllTestimonials(),
        excelDb.getAllSubscribers(),
        excelDb.getAllBlogs()
      ]);

      const totalRevenue = orders.reduce((sum, o) => sum + (o.total || 0), 0);
      const paidEnrollmentsCount = enrollments.filter(e => e.paymentStatus === 'Paid').length;

      res.json({
        totalUsers: users.length,
        totalCourses: courses.length,
        totalEnrollments: enrollments.length,
        paidEnrollments: paidEnrollmentsCount,
        totalContacts: contacts.length,
        totalAppointments: appointments.length,
        totalOrders: orders.length,
        totalBlogs: blogs.length,
        totalRevenue: Math.round(totalRevenue * 100) / 100,
        totalTestimonials: testimonials.length,
        totalSubscribers: subscribers.length,
        recentEnrollments: enrollments.slice(-5).reverse(),
        recentContacts: contacts.slice(-5).reverse(),
        recentAppointments: appointments.slice(-5).reverse(),
        recentOrders: orders.slice(-5).reverse(),
        courseDistribution: courses.map(c => ({
          category: c.category,
          title: c.title,
          enrollmentsCount: enrollments.filter(e => e.courseId === c.id).length
        }))
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // --- VITE MIDDLEWARE SETUP ---
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[Evolutionary Web Dude Server] Running on http://localhost:${PORT}`);
    console.log(`[Storage] Native Excel sheet storage initialized: ${excelDb.getFilePath()}`);
  });
}

startServer();

const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// =========================
// CONFIG
// =========================
const ADMIN_PASS = process.env.ADMIN_PASSWORD || '';
const SHEET_ID = process.env.GOOGLE_SHEET_ID;

const SHEET_SCHEMAS = {
  Users: ['phone', 'name', 'email', 'avatar', 'created_at'],
  Products: ['id', 'name', 'price', 'desc', 'img'],
  Orders: ['id', 'time', 'customer', 'phone', 'items', 'total', 'status']
};

// =========================
// GOOGLE SHEETS CORE
// =========================
function getAuth() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);

  return new JWT({
    email: creds.client_email,
    key: creds.private_key.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getDoc() {
  const doc = new GoogleSpreadsheet(SHEET_ID, getAuth());
  await doc.loadInfo();
  return doc;
}

async function getSheet(doc, title) {
  let sheet = doc.sheetsByTitle[title];

  if (!sheet) {
    sheet = await doc.addSheet({
      title,
      headerValues: SHEET_SCHEMAS[title] || ['id'],
    });
  } else {
    try {
      await sheet.loadHeaderRow();
    } catch {
      await sheet.setHeaderRow(SHEET_SCHEMAS[title]);
    }
  }

  return sheet;
}

// =========================
// USERS (UPSERT)
// =========================
app.post('/api/user/login', async (req, res) => {
  try {
    const { phone, name, email, avatar } = req.body;

    if (!phone || phone.length !== 11) {
      return res.status(400).json({ error: 'Phone must be 11 digits' });
    }

    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Users');

    const rows = await sheet.getRows();
    const existing = rows.find(r => String(r.phone) === String(phone));

    if (existing) {
      existing.name = name || existing.name;
      existing.email = email || existing.email;
      existing.avatar = avatar || existing.avatar;
      await existing.save();

      return res.json({
        exists: true,
        user: { phone, name: existing.name, email: existing.email, avatar: existing.avatar }
      });
    }

    await sheet.addRow({
      phone,
      name: name || 'User',
      email: email || '',
      avatar: avatar || '',
      created_at: new Date().toISOString()
    });

    res.json({
      exists: false,
      user: { phone, name, email, avatar }
    });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =========================
// PRODUCTS
// =========================
app.get('/api/products', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Products');

    const rows = await sheet.getRows();

    const products = rows.map(r => ({
      id: String(r.id),
      name: r.name,
      price: Number(r.price),
      desc: r.desc,
      img: r.img
    })).filter(p => p.name);

    res.json(products);

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/products/save', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Products');

    const product = {
      id: String(req.body.id || Date.now()),
      name: req.body.name || '',
      price: Number(req.body.price) || 0,
      desc: req.body.desc || '',
      img: req.body.img || ''
    };

    const rows = await sheet.getRows();
    const existing = rows.find(r => String(r.id) === product.id);

    if (existing) {
      existing.name = product.name;
      existing.price = product.price;
      existing.desc = product.desc;
      existing.img = product.img;
      await existing.save();

      return res.json({ ok: true, action: 'updated' });
    }

    await sheet.addRow(product);

    res.json({ ok: true, action: 'created' });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/products/delete', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Products');

    const rows = await sheet.getRows();
    const row = rows.find(r => String(r.id) === String(req.body.id));

    if (!row) {
      return res.status(404).json({ error: 'Product not found' });
    }

    await row.delete();

    res.json({ ok: true });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =========================
// ADMIN SAVE ALL PRODUCTS
// =========================
app.post('/api/admin/save', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Products');

    const products = req.body.products || [];

    await sheet.clear();
    await sheet.setHeaderRow(SHEET_SCHEMAS.Products);

    if (products.length) {
      await sheet.addRows(products);
    }

    res.json({ ok: true, count: products.length });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =========================
// ORDERS
// =========================
app.post('/api/order', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Orders');

    const { customer, phone, items, total } = req.body;

    const orderId = 'ORD' + Date.now();

    const itemsText = (items || [])
      .map(i => `${i.name} x1 - ₦${Number(i.price).toLocaleString()}`)
      .join(', ');

    await sheet.addRow({
      id: orderId,
      time: new Date().toISOString(),
      customer: customer || 'Guest',
      phone: phone || '',
      items: itemsText,
      total: Number(total) || 0,
      status: 'pending'
    });

    res.json({ success: true, orderId });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/orders', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Orders');

    const rows = await sheet.getRows();

    const orders = rows.map(r => ({
      id: r.id,
      time: r.time,
      customer: r.customer,
      phone: r.phone,
      items: r.items,
      total: Number(r.total),
      status: r.status
    })).reverse();

    res.json(orders);

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/orders/update', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Orders');

    const rows = await sheet.getRows();
    const row = rows.find(r => r.id === req.body.id);

    if (!row) return res.status(404).json({ error: 'Order not found' });

    row.status = req.body.status;
    await row.save();

    res.json({ ok: true });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/orders/delete', async (req, res) => {
  try {
    const doc = await getDoc();
    const sheet = await getSheet(doc, 'Orders');

    const rows = await sheet.getRows();
    const row = rows.find(r => r.id === req.body.id);

    if (!row) return res.status(404).json({ error: 'Order not found' });

    await row.delete();

    res.json({ ok: true });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =========================
// ADMIN AUTH
// =========================
app.post('/api/admin/auth', (req, res) => {
  if (!ADMIN_PASS) {
    return res.status(403).json({ error: 'No admin password set' });
  }

  if (req.body.password === ADMIN_PASS) {
    res.json({ ok: true });
  } else {
    res.status(401).json({ error: 'Wrong password' });
  }
});

// =========================
// DEBUG
// =========================
app.get('/api/debug', (req, res) => {
  res.json({
    sheet: SHEET_ID ? 'SET' : 'MISSING',
    creds: process.env.GOOGLE_CREDENTIALS ? 'SET' : 'MISSING',
    admin: ADMIN_PASS ? 'SET' : 'MISSING'
  });
});

// =========================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
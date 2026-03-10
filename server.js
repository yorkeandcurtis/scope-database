const express = require('express');
const basicAuth = require('express-basic-auth');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(basicAuth({
  users: { [process.env.AUTH_USER || 'yc']: process.env.AUTH_PASS || 'changeme' },
  challenge: true,
  realm: 'Yorke & Curtis Scope Database'
}));

app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => {
  console.log(`Scope database running on port ${PORT}`);
});

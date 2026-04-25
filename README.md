# A VAN CANDIES

A candy-themed website for A VAN CANDIES with:

- a colorful landing page
- a menu page section with current prices
- a separate order page
- a local Python order server that appends submitted orders into a private Excel file on the same computer

## Files

- `index.html` - main landing page
- `styles.css` - main site styles
- `order.html` - order form page
- `order.css` - order page styles
- `order.js` - sends order data to the local server
- `order_server.py` - local server that saves orders into a private workbook
- `start_order_server.bat` - quick starter for the local order server

## Local Order Saving

The order page is designed to work with the local Python server.

1. Start the server:

```bat
start_order_server.bat
```

2. Open:

```text
http://127.0.0.1:8001/order.html
```

3. Submit an order. The server appends the order into a private Excel file on the local machine.

## Privacy

- The server blocks direct `.xlsx` downloads over the website.
- Customer order files are not meant to be committed to GitHub.
- Orders are intended to stay on the local machine running the server.

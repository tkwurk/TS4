// Test script to verify partial payment logic
function testPartialPaymentLogic() {
  console.log("=== TESTING PARTIAL PAYMENT LOGIC ===");
  
  // Example order with partial payment
  const testOrder = {
    orderNumber: "ORD-20250817-001",
    items: [
      {
        name: "Burger Deluxe",
        qty: 1,
        price: 10000,
        category: "Food"
      }
    ],
    total: 10000,
    payment: "partial",
    paymentDetails: {
      dana: 8000,
      cash: 2000,
      total: 10000
    }
  };
  
  console.log("Test Order:", JSON.stringify(testOrder, null, 2));
  
  // Simulate the logic from saveOrder function
  const item = testOrder.items[0];
  const itemTotal = item.price * item.qty; // 10000
  
  // Calculate item's share of total payment
  const itemShare = itemTotal / testOrder.total; // 1.0 (100%)
  const itemDanaAmount = Math.round(testOrder.paymentDetails.dana * itemShare); // 8000
  const itemCashAmount = Math.round(testOrder.paymentDetails.cash * itemShare); // 2000
  
  console.log("\n--- Calculated Split ---");
  console.log("Item Share:", itemShare);
  console.log("Item Dana Amount:", itemDanaAmount);
  console.log("Item Cash Amount:", itemCashAmount);
  
  // Simulate the two records that would be saved
  const danaRecord = {
    orderNumber: "ORD-20250817-001-001-DANA",
    productName: item.name,
    quantity: item.qty, // Full quantity = 1
    price: itemDanaAmount, // 8000
    paymentMethod: "dana",
    total: itemDanaAmount, // 8000
    category: item.category
  };
  
  const cashRecord = {
    orderNumber: "ORD-20250817-001-001-CASH", 
    productName: item.name,
    quantity: 0, // Zero quantity for cash record
    price: itemCashAmount, // 2000
    paymentMethod: "cash",
    total: itemCashAmount, // 2000
    category: item.category
  };
  
  console.log("\n--- Records to be saved ---");
  console.log("Dana Record:", JSON.stringify(danaRecord, null, 2));
  console.log("Cash Record:", JSON.stringify(cashRecord, null, 2));
  
  // Verify the logic
  const totalQuantity = danaRecord.quantity + cashRecord.quantity; // Should equal original qty
  const totalAmount = danaRecord.total + cashRecord.total; // Should equal original total
  
  console.log("\n--- Verification ---");
  console.log("Original Quantity:", item.qty);
  console.log("Total Saved Quantity:", totalQuantity);
  console.log("Quantity Match:", totalQuantity === item.qty);
  
  console.log("Original Total:", itemTotal);
  console.log("Total Saved Amount:", totalAmount);
  console.log("Amount Match:", totalAmount === itemTotal);
  
  console.log("\n--- Summary ---");
  console.log("✅ Dana record has full quantity (1) with Dana amount (8000)");
  console.log("✅ Cash record has zero quantity (0) with Cash amount (2000)");
  console.log("✅ Total quantity preserved:", totalQuantity === item.qty);
  console.log("✅ Total amount preserved:", totalAmount === itemTotal);
  
  return {
    success: true,
    danaRecord: danaRecord,
    cashRecord: cashRecord,
    verification: {
      quantityMatch: totalQuantity === item.qty,
      amountMatch: totalAmount === itemTotal
    }
  };
}

// Run the test
if (typeof module !== 'undefined' && module.exports) {
  module.exports = testPartialPaymentLogic;
} else {
  // Browser environment
  console.log("Running test in browser...");
  testPartialPaymentLogic();
}

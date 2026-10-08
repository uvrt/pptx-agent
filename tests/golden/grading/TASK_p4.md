# Task: draw a process diagram

`input/order-to-cash.pptx` has an empty slide 2, "Order-to-cash process". Draw the process on it as a flow chart:

1. **Order received**
2. **Credit check**: a decision. If the credit check fails, the flow goes to **Notify customer**, which ends there. If it passes, the flow continues:
3. **Pick and pack**
4. **Ship order**
5. **Send invoice**
6. **Payment received**

Requirements:

- Each step is a shape with its text in it. The credit check is a decision (diamond) shape; the other steps are rectangles or rounded rectangles.
- The main flow reads left to right (it may wrap onto a second row); "Notify customer" sits off the main flow, near the credit check.
- Every arrow is a **connector attached** to the two shapes it joins (so it follows them if they are moved), with an arrowhead pointing in the direction of the flow. Label the two arrows leaving the credit check "Pass" and "Fail" (a small text box next to each arrow is fine).
- Use the **theme's colours**, not hard-coded RGB: one accent colour for ordinary steps, a different one for the decision, and a third for "Notify customer". Text must be readable on its fill.
- Everything stays inside the slide and below the title; no shapes overlap each other, and no text overflows its shape.
- The other slides do not change.

Save the result as `work/order-to-cash.pptx`.

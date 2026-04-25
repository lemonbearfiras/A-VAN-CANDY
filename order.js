(function () {
    const form = document.getElementById("orderForm");
    const status = document.getElementById("formStatus");
    const LOCAL_SERVER_ORIGIN = "http://127.0.0.1:8001";
    const LOCAL_SAVE_ENDPOINT = LOCAL_SERVER_ORIGIN + "/submit-order";

    if (!form || !status) {
        return;
    }

    const submitButton = form.querySelector('button[type="submit"]');

    form.addEventListener("submit", async function (event) {
        event.preventDefault();

        const formData = new FormData(form);
        const order = {
            name: String(formData.get("name") || "").trim(),
            email: String(formData.get("email") || "").trim(),
            class: String(formData.get("class") || "").trim(),
            school: String(formData.get("school") || "").trim(),
            phone: String(formData.get("phone") || "").trim(),
            item: String(formData.get("item") || "").trim(),
            quantity: String(formData.get("quantity") || "").trim(),
            notes: String(formData.get("notes") || "").trim(),
        };

        if (!order.name || !order.email || !order.class || !order.school || !order.phone || !order.item || !order.quantity) {
            setStatus("Please complete all required fields before submitting.", "error");
            return;
        }

        try {
            if (submitButton) {
                submitButton.disabled = true;
                submitButton.textContent = "Saving...";
            }

            const response = await fetch(LOCAL_SAVE_ENDPOINT, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify(order),
            });

            const result = await response.json().catch(function () {
                return {};
            });

            if (!response.ok) {
                throw new Error(result.error || "Could not save the order.");
            }

            setStatus("Order saved into the Excel file.", "success");
            form.reset();
            form.elements.quantity.value = "1";
        } catch (error) {
            setStatus(error.message || "Could not save the order. Make sure the local server is running on this computer.", "error");
        } finally {
            if (submitButton) {
                submitButton.disabled = false;
                submitButton.textContent = "Submit Order";
            }
        }
    });

    function setStatus(message, state) {
        status.textContent = message;
        status.dataset.state = state;
    }
}());

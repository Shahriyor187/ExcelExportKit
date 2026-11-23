const $ = (selector) => document.querySelector(selector);
const feedback = $("#feedback");
const money = new Intl.NumberFormat("en-US", {style:"currency",currency:"USD",maximumFractionDigits:0});

async function loadPreview() {
  try {
    const response = await fetch("/api/preview");
    if (!response.ok) throw new Error("Unable to load the sample dataset");
    const orders = await response.json();
    $("#record-count").textContent = String(orders.length);
    const body = $("#orders-body");
    body.replaceChildren();
    for (const order of orders) {
      const tr = document.createElement("tr");
      for (const value of [
        order.orderId, order.customer, order.region, order.product,
        order.units, money.format(order.revenue)
      ]) {
        const td = document.createElement("td");
        td.textContent = String(value);
        tr.append(td);
      }
      body.append(tr);
    }
  } catch (error) {
    $("#orders-body").replaceChildren();
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 6;
    td.textContent = "The sample data could not be loaded. Check that the app is running.";
    tr.append(td);
    $("#orders-body").append(tr);
  }
}

async function downloadExport(button) {
  const mode = button.dataset.mode;
  const format = button.dataset.format;
  button.disabled = true;
  feedback.classList.remove("error");
  feedback.textContent = "Generating " + format.toUpperCase() + " export…";
  try {
    const response = await fetch("/api/export/" + mode + "/" + format);
    if (!response.ok) throw new Error("The server rejected this export (" + response.status + ").");
    const blob = await response.blob();
    if (!blob.size) throw new Error("The exported file was empty.");
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "export-studio-" + mode + "." + format;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    feedback.textContent = "Downloaded " + format.toUpperCase() + " successfully (" + (blob.size / 1024).toFixed(1) + " KB).";
  } catch (error) {
    feedback.classList.add("error");
    feedback.textContent = error.message || "The export failed.";
  } finally {
    button.disabled = false;
  }
}

document.querySelectorAll("[data-mode][data-format]").forEach((button) => {
  button.addEventListener("click", () => downloadExport(button));
});
loadPreview();

const express = require("express");
const router = express.Router();

router.get("/", (req, res) => {
    res.render("pages/home_2");
  });
router.get("/email", (req, res) => {
    res.render("pages/email");
  });



module.exports = router;
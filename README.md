## Summary

<!--
**YaoqianM/YaoqianM** is a ✨ _special_ ✨ repository because its `README.md` (this file) appears on your GitHub profile.

Here are some ideas to get you started:


- 🌱 I’m currently learning ...
- 👯 I’m looking to collaborate on ...
- 🤔 I’m looking for help with ...
- 💬 Ask me about ...
- 📫 How to reach me: ...
- 😄 Pronouns: ...
-  Fun fact: ...
-->

⚡B.S. -> Computer Science, 
Master-> Engineering Management && Information Systems Management

⚡Project Experience: 
Gaming Web Platform, Video-Sharing App, Distributed KV Storage System and E-Commerce Platform.

💬Java, C++, JavaScript, Kotlin, TypeScript, HTML, Redis, MySQL, MongoDB, Oracle SQL\
💬SpringBoot, Spring, React, MyBatis, Node.js, RocketMQ, Zookeeper, AWS Cloud\
💬Docker, Kafka, JVM, VMware, Git, Maven, IntelliJ, Visual Studio Code, Agile, DevOps, ITIL

## Mianshiya question scraper

Run the script below to export every question from [面试鸭](https://www.mianshiya.com/) into an Excel workbook (only the question texts are stored):

```bash
pip install -r requirements.txt
python mianshiya_scraper.py --output mianshiya_questions.xlsx
```

To generate a one-click executable for Windows 11, install [PyInstaller](https://pyinstaller.org/) and run:

```bash
pyinstaller --onefile mianshiya_scraper.py
```

The resulting `dist/mianshiya_scraper.exe` can be double-clicked to export the questions.  Place the executable next to the generated Excel file or wrap it in a `.bat` script to supply custom arguments.

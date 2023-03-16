const puppeteer = require('puppeteer');
const fs = require("fs")
const parser = new (require('simple-excel-to-json').XlsParser)();

(async function () {
    try {
        const raws = parser.parseXls2Json('./messages.xlsx');
        const sample = fs.readFileSync("./sample.txt", { encoding: 'utf-8' });
        console.log(sample)
        const messages = raws[0].map(raw => {
            console.log(raw)
            let phoneNumber = raw.phone.toString();
            if (!phoneNumber.startsWith("0")) phoneNumber = `0${phoneNumber}`
            delete raw.phone;
            let content = sample
            for (const [key, value] of Object.entries(raw)) {
                const regexStr = `#{${key}}`
                const regex = new RegExp(regexStr, 'g')
                content = content.replace(regex, value)
            }
            return { phoneNumber, content }
        })
        console.log(messages)

        const browser = await puppeteer.launch({ executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', headless: false });
        const page = await browser.newPage();
        await page.setViewport({ width: 1200, height: 700 });

        await page.goto(`https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F`, { waitUntil: 'networkidle2' })

        const avatarSelector = ".zavatar.zavatar-l.zavatar-single.flx.flx-al-c.flx-center.rel.clickable.disableDrag";
        await page.waitForSelector(avatarSelector);

        for (const message of messages) {
            try {
                await page.goto(`https://chat.zalo.me/?phone=${message.phoneNumber}`, { waitUntil: 'networkidle2' })
                const sendMsgBtnSelector = ".btnpf-content"
                await page.waitForSelector(sendMsgBtnSelector);
                await page.click(sendMsgBtnSelector)

                const chatInputSelector = "#chatInput"
                await page.waitForSelector(chatInputSelector)
                await page.type(chatInputSelector, message.content)
                await page.keyboard.press('Enter')

                const sendTxtSelector = 'span[data-translate-inner="STR_SENT"]'
                await page.waitForSelector(sendTxtSelector)
                console.log(`Sent to  ${message.phoneNumber}`)
            } catch (error) {
                console.error(error)
                console.error(`Error with ${message.phoneNumber}`)
            }
        }
    } catch (error) {
        console.error(error)
    }
})()
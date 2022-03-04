const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx');
const tesseract = require('tesseract.js');
const { createWorker } = require('tesseract.js');

// const getCode = async src => {
//     const data = await tesseract.recognize(src);
//     if (data) {
//         return data.data.text;
//     }
//     return false;
// };
const getCode = async src => {
    try {
        const worker = createWorker();
        await worker.load();
        await worker.loadLanguage('eng');
        await worker.initialize('eng');
        const {
            data: { text }
        } = await worker.recognize(src);
        return text;
    } catch (err) {
        return false;
    }
};

const workSheetsFromFile = xlsx.parse(`${__dirname}/工作表.xlsx`);
const getAccountArr = () => {
    const arr = [];
    for (const item of workSheetsFromFile) {
        if (item.name == 'Sheet1') {
            const data = item.data;
            const len = data.length;
            for (let i = 1; i < len; i++) {
                const child = data[i];
                arr.push({
                    id: child[3].toString(),
                    name: child[1],
                    phone: child[4].toString()
                });
            }
        }
    }
    console.log(arr);
    return arr;
};
// const accountArr = getAccountArr();
const accountArr = [
    {
        id: '420902199406122253',
        name: '张三',
        phone: 'Aa12345678'
    }
];
const pageConfig = {
    headless: true,
    args: ['--start-fullscreen'],
    defaultViewport: {
        width: 1920,
        height: 1080
    }
};
const pageUrl = 'http://person.zhujianpeixun.com/Login/Login';
const delay = 100;

const getScreenshot = async (page, codeElBox, codeImgName) => {
    const { x, y, width, height } = codeElBox;
    await page.screenshot({
        path: `./code/${codeImgName}.png`,
        clip: {
            x,
            y,
            width,
            height
        }
    });
};

let total = accountArr.length;
let successArr = [];
let errorArr = [];

const wait = async time => {
    return new Promise(resolve => {
        setTimeout(() => {
            resolve(true);
        }, time || 2000);
    });
};

// 启动
const launchFun = async () => {
    console.time('本次运行耗时');
    console.log('启动中...');
    console.log('开始自动操作...');
    for (const item of accountArr) {
        const browser = await puppeteer.launch(pageConfig);
        await wait(4000);
        optionFun(item, browser, true); // 第三个参数表示改密码 还是验证结果
    }
};

launchFun();

const isFinished = async (browser, page) => {
    page.removeListener('response');
    await browser.close();
    const succLen = successArr.length;
    const errLen = errorArr.length;
    if (succLen + errLen >= total) {
        console.log('操作完毕');
        console.log(`总共${total}个用户`);
        console.log(`成功${succLen}个`);
        console.log(`失败${errLen}个`);
        if (succLen > 0) {
            console.log('————————————————————————————');
            console.log('成功用户：');
            for (const item of successArr) {
                console.log(`姓名：${item.name}`);
                console.log(`身份证号：${item.id}`);
            }
        }
        if (errLen > 0) {
            console.log('————————————————————————————');
            console.log('失败用户：');
            for (const item of errorArr) {
                console.log(`姓名：${item.name}`);
                console.log(`身份证号：${item.id}`);
                item.reason && console.log(`失败原因：${item.reason}`);
            }
        }
        console.timeEnd('本次运行耗时');
    }
};

// 执行过程
const optionFun = async (item, browser, validate) => {
    const page = await browser.newPage();
    await page.goto(pageUrl);

    let reTryTime = 0;
    process.setMaxListeners(0);
    page.on('response', async response => {
        const responseUrl = response.url();
        // 这个接口成功后页面会重定向，导致拿不到response，会走catch
        if (responseUrl.indexOf('/Login/AdminLogin/') > -1) {
            try {
                const res = await response.json();
                if (!res.success) {
                    if (res.error.includes('用户名或密码错误')) {
                        console.log(
                            `用户 ${item.name} 登录失败，用户名或密码错误`
                        );
                        errorArr.push({
                            ...item,
                            reason: '用户名或密码错误'
                        });
                        await isFinished(browser, page);
                        return;
                    }
                    reTryTime++;
                    if (reTryTime <= 5) {
                        console.log(
                            `用户 ${item.name} 登录失败，正在进行第${reTryTime}次重试`
                        );
                        await typeCode(page, item.id);
                    } else {
                        console.log(`用户 ${item.name} 登录失败，验证码错误`);
                        errorArr.push({ ...item, reason: '验证码错误' });
                        await isFinished(browser, page);
                    }
                } else {
                    if (!validate) {
                        console.log(
                            `用户 ${item.name} 登录成功，正在修改密码...`
                        );
                        await changePwd(page, item);

                        successArr.push(item);
                        await isFinished(browser, page);
                    }
                }
            } catch (err) {
                if (validate) {
                    // 走catch证明页面重定向了，表示登录成功，不需要修改密码
                    console.log(`用户 ${item.name} 登录成功`);
                    successArr.push(item);
                    await isFinished(browser, page);
                }
            }
        }
    });

    await loginFun(page, item, validate);
};

// 改密码
const changePwd = async (page, item) => {
    await page.waitForSelector('.layui-layer-btn0');
    const confirmChangeElement = await page.$('.layui-layer-btn0');
    await confirmChangeElement.click();

    const UserAccount = await page.$('#UserAccount');
    await UserAccount.type(item.id, { delay });

    const Password_UpdatePwd = await page.$('#Password_UpdatePwd');
    await Password_UpdatePwd.type(item.phone, { delay });

    const NPassword_UpdatePwd = await page.$('#NPassword_UpdatePwd');
    await NPassword_UpdatePwd.type('Aa12345678', { delay });

    const TruePassword_UpdatePwd = await page.$('#TruePassword_UpdatePwd');
    await TruePassword_UpdatePwd.type('Aa12345678', { delay });

    const updateElement = await page.$('button[lay-filter=UpdatePwd]');
    await updateElement.click();
};

// 登录
const loginFun = async (page, item, validate) => {
    await page.waitForSelector('input[value=我知道了]');
    const iKnowElement = await page.$('input[value=我知道了]');
    await iKnowElement.click();

    const loginBtn = await page.$('.quick-l');
    await loginBtn.click();

    const user_name_student = await page.$('#user_name_student');
    await user_name_student.type(item.id, { delay });

    const user_pwd_student = await page.$('#user_pwd_student');
    if (validate) {
        await user_pwd_student.type('Aa12345678', { delay });
    } else {
        await user_pwd_student.type(item.phone, { delay });
    }
    await typeCode(page, item.id);
};

// 输入验证码
const typeCode = async (page, codeImgName) => {
    try {
        const imgCode_student = await page.$('#imgCode_student');
        const codeElBox = await imgCode_student.boundingBox();
        await imgCode_student.click();
        await page.waitForResponse(
            response =>
                response.url().indexOf('/Login/CheckCode?') > -1 &&
                response.status() === 200
        );
        await getScreenshot(page, codeElBox, codeImgName);
        const code = await getCode(`./code/${codeImgName}.png`);
        if (!code || !(code - 0)) {
            await typeCode(page, codeImgName);
        } else {
            const Code_student = await page.$('#Code_student');
            await Code_student.click({ clickCount: 3 });
            await Code_student.type(code, { delay });
        }
    } catch (err) {}
};

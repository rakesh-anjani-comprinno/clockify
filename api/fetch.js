import axios from 'axios'
import dotenv from 'dotenv'
dotenv.config()
const BASE_URL = "https://api.clockify.me/";
const clockifyApiKey = process.env.API_KEY
const headers = {
    'X-Api-Key': clockifyApiKey
};


const maxRetries = 50;


export const fetchDataFromApi = async (url, params, currentRetry = 0) => {
    let results = []
    let page = 1
    while (true) {
        try {
            const response = await axios.get(BASE_URL + url, {
                headers,
                params: { ...params, page: page },
            });
            const responseData = response.data
            results = results.concat(responseData)
            if (responseData.length < params.page_size) {
                break
            }
            page += 1;
        } catch (err) {
            if (err.response.status === 429) {
                const retryAfter = err.response.headers.get('Retry-After') || 1;
                await wait(retryAfter * 1000);
                if (currentRetry < maxRetries) {
                    currentRetry++;
                    // console.log("currentretry",url,params,"page**",page,currentRetry)
                } else {
                    console.log("Max retries reached", page, params, url)
                    return err.response.data
                }
            } else {
                console.log("response error data", err.response.data)
                return err.response.data
            }
        }
    }
    return results
};

const wait = (ms) => {
    return new Promise(resolve => setTimeout(resolve, ms));
}
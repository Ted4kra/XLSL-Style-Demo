'use strict'
const path = require('path')

function resolve(dir) {
    return path.join(__dirname, dir)
}

module.exports={
    configureWebpack: {
        externals: {
            'XLSX': 'XLSX'
        }
    },

    chainWebpack:(config)=>{
        config.resolve.alias
            //set第一个参数：设置的别名，第二个参数：设置的路径
            .set('@',resolve('./src'))
            .set('components',resolve('./src/components'))
            .set('assets',resolve('./src/assets'))
            .set('sources',resolve('./src/sources'))
            .set('views',resolve('./src/views'))
    }
}

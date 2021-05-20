import FileSaver from 'file-saver'
import xlsx from 'XLSX'

class ExportReportExcel {
  constructor(target) {
    this.target = target
  }

  transformD(data) {
    // 表头信息
    let aoa = [
      ['附件', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
      ['江苏省多式联运示范工程项目' + (this.target.excelYear) + this.target.excelQuarter + '动态监测信息表'],
      ['          填报周期：_______年____季度                               填报人：                                   联系电话：                                   '],
      ['序号', '市别', '示范工程项目名称', '示范工程项目运行情况（累计至本季度）'],
      [null, null, null, '联运线路（条）', '线路', '联运线路', '联运模式', '示范线路长度', null, null, null, '联运量', null, '联运周转量（万吨公里）', '平均运转次数（次）', '运输主要货品货类', '平均单位联运价格（元/吨﹒公里）', '平均公路运输价格（元/吨﹒公里）', '多式联运量占比（%）', '集装箱多式联运量占比%', '企业同期总货物运输量', null, '联运设施设备', null, null, null, null, null, '联运信息化', null, null, '其他'],
      [null, null, null, null, null, null, null, '总长度（公里）', '公路', '铁路', '水路', '万吨', '万标箱', null, null, null, null, null, null, null, '万吨', '万标箱', '枢纽场站建设季度投资完成额（万元）', '枢纽场站建设累计完成投资额（万元）', '枢纽场站投资完成率（%）', '装备设备购置季度投资完成额（万元）', '装备设备购置累计完成投资额（万元）', '装备设备购置完成率（%）', '信息系统建设季度投资完成额（万元）', '信息系统建设累计完成投资额（万元）', '信息系统建设完成率（%）', null]
    ]
    // 表头合并
    const mergeTitle = [
      // 江苏省多式联运示范工程项目2019年第二季度动态监测信息表
      {
        s: {
          r: 1,
          c: 0
        },
        e: {
          r: 1,
          c: 29
        }
      },
      // '填报周期：_______年____季度                               填报人：                                   联系电话：                                   \'
      {
        s: {
          r: 2,
          c: 0
        },
        e: {
          r: 2,
          c: 29
        }
      },
      // 序号
      {
        s: {
          r: 3,
          c: 0
        },
        e: {
          r: 5,
          c: 0
        }
      },
      // 市别
      {
        s: {
          r: 3,
          c: 1
        },
        e: {
          r: 5,
          c: 1
        }
      },
      // 示范工程名称
      {
        s: {
          r: 3,
          c: 2
        },
        e: {
          r: 5,
          c: 2
        }
      },
      // 示范工程项目运行情况
      {
        s: {
          r: 3,
          c: 3
        },
        e: {
          r: 3,
          c: 29
        }
      },
      // 线路总条数
      {
        s: {
          r: 4,
          c: 3
        },
        e: {
          r: 5,
          c: 3
        }
      },
      // 线路
      {
        s: {
          r: 4,
          c: 4
        },
        e: {
          r: 5,
          c: 4
        }
      },
      // 联运路线
      {
        s: {
          r: 4,
          c: 5
        },
        e: {
          r: 5,
          c: 5
        }
      },
      // 联运模式
      {
        s: {
          r: 4,
          c: 6
        },
        e: {
          r: 5,
          c: 6
        }
      },
      // 示范线路长度
      {
        s: {
          r: 4,
          c: 7
        },
        e: {
          r: 4,
          c: 10
        }
      },
      // 总长度
      {
        s: {
          r: 5,
          c: 7
        },
        e: {
          r: 5,
          c: 7
        }
      },
      // 公路
      {
        s: {
          r: 5,
          c: 8
        },
        e: {
          r: 5,
          c: 8
        }
      },
      // 铁路
      {
        s: {
          r: 5,
          c: 9
        },
        e: {
          r: 5,
          c: 9
        }
      },
      // 水路
      {
        s: {
          r: 5,
          c: 10
        },
        e: {
          r: 5,
          c: 10
        }
      },
      // 联运量
      {
        s: {
          r: 4,
          c: 11
        },
        e: {
          r: 4,
          c: 12
        }
      },
      // 联运量-万吨
      {
        s: {
          r: 5,
          c: 11
        },
        e: {
          r: 5,
          c: 11
        }
      },
      // 联运量-万标箱
      {
        s: {
          r: 5,
          c: 12
        },
        e: {
          r: 5,
          c: 12
        }
      },
      // 联运周转量
      {
        s: {
          r: 4,
          c: 13
        },
        e: {
          r: 5,
          c: 13
        }
      },
      // 平均运转次数
      {
        s: {
          r: 4,
          c: 14
        },
        e: {
          r: 5,
          c: 14
        }
      },
      // 运输主要货品货类
      {
        s: {
          r: 4,
          c: 15
        },
        e: {
          r: 5,
          c: 15
        }
      },
      // 平均单位联运价格
      {
        s: {
          r: 4,
          c: 16
        },
        e: {
          r: 5,
          c: 16
        }
      },
      // 平均公路运输价格
      {
        s: {
          r: 4,
          c: 17
        },
        e: {
          r: 5,
          c: 17
        }
      },
      // 多式联运量占比%
      {
        s: {
          r: 4,
          c: 18
        },
        e: {
          r: 5,
          c: 18
        }
      },
      // 集装箱多式联运量占比%
      {
        s: {
          r: 4,
          c: 19
        },
        e: {
          r: 5,
          c: 19
        }
      },
      // 企业总货物运输量
      {
        s: {
          r: 4,
          c: 20
        },
        e: {
          r: 4,
          c: 21
        }
      },
      // 企业同期总货物运输量-万吨
      {
        s: {
          r: 5,
          c: 20
        },
        e: {
          r: 5,
          c: 20
        }
      },
      // 企业同期总货物运输量-万标箱
      {
        s: {
          r: 5,
          c: 21
        },
        e: {
          r: 5,
          c: 21
        }
      },
      // 联运设施设备
      {
        s: {
          r: 4,
          c: 22
        },
        e: {
          r: 4,
          c: 27
        }
      },
      // 枢纽场站建设季度投资完成额（万元）
      {
        s: {
          r: 5,
          c: 22
        },
        e: {
          r: 5,
          c: 22
        }
      },
      // 枢纽场站建设累计完成投资额
      {
        s: {
          r: 5,
          c: 23
        },
        e: {
          r: 5,
          c: 23
        }
      },
      // 枢纽场站建设投资完成率
      {
        s: {
          r: 5,
          c: 24
        },
        e: {
          r: 5,
          c: 24
        }
      },
      // 装备设备购置季度投资完成额（万元）
      {
        s: {
          r: 5,
          c: 25
        },
        e: {
          r: 5,
          c: 25
        }
      },
      // 装备设备累计完成购置投资额
      {
        s: {
          r: 5,
          c: 26
        },
        e: {
          r: 5,
          c: 26
        }
      },
      // 装备设备购置投资完成率
      {
        s: {
          r: 5,
          c: 27
        },
        e: {
          r: 5,
          c: 27
        }
      },
      // 联运信息化
      {
        s: {
          r: 4,
          c: 28
        },
        e: {
          r: 4,
          c: 30
        }
      },
      // 信息系统建设季度投资完成额（万元）
      {
        s: {
          r: 5,
          c: 28
        },
        e: {
          r: 5,
          c: 28
        }
      },
      // 信息系统建设累计完成投资额
      {
        s: {
          r: 5,
          c: 29
        },
        e: {
          r: 5,
          c: 29
        }
      },
      // 信息系统建设投资完成率
      {
        s: {
          r: 5,
          c: 30
        },
        e: {
          r: 5,
          c: 30
        }
      },
      // 其他
      {
        s: {
          r: 4,
          c: 31
        },
        e: {
          r: 5,
          c: 31
        }
      }
    ]
    // 数据
    const requestResult = data
    const totalForm = []
    const mergeContent = []
    let startMergeLength = aoa.length
    // 遍历主表数据
    /*
        [{
      "id": "b8e0cc2f-8ea1-45aa-9e3d-3f93900b6934",
      "projectId": "2a7a26c0-42a0-444a-ba56-e60b3d41a667",
      "year": 2020,
      "quarter": 1,
      "hubStationBuildInvestMoney": 12.0,
      "hubStationBuildInvestRate": 12.0,
      "equipmentInvestMoney": 12.0,
      "equipmentInvestRate": 12.0,
      "informationSysInvestMoney": 12.0,
      "informationSysInvestRate": 12.0,
      "remark": "12",
      "creator": null,
      "createTime": null,
      "modifier": null,
      "modifyTime": null,
      "auditStatusId": "PENDING_AUDIT",
      "cityId": "320100",
      "cityName": "南京市",
      "projectName": "有两条线路",
      "detailList": [{
        "id": "792b0f17-55d6-4e69-8d10-efba58ecc11e",
        "reportId": "b8e0cc2f-8ea1-45aa-9e3d-3f93900b6934",
        "lineId": "0bc80db1-709c-40b7-868f-43e0e1fcc38d",
        "transportWeightWd": 11.0,
        "transportContainerCount": 22.0,
        "turnWeight": 33.0,
        "averageNum": 44.0,
        "cargoType": null,
        "averagePriceHighway": 1.0,
        "averageRoadTransportPrice": 2.0,
        "companyTotalWeightWd": 3.0,
        "companyTotalWeightWCt": 4.0,
        "remark": "备注信息",
        "creator": null,
        "createTime": null,
        "modifier": null,
        "modifyTime": null,
        "transportLine": "xxxxxxxxxxx",
        "transportModeCode": "ROAD_WATER",
        "totalLength": 123.12,
        "roadLength": 1111.0,
        "railwayLength": 2222.0,
        "waterLength": 3333.0
      }, {
        "id": "2286c9f3-e1ef-4887-928c-2d95d0a06b5d",
        "reportId": "b8e0cc2f-8ea1-45aa-9e3d-3f93900b6934",
        "lineId": "a8519c52-1e92-4c98-b196-76cb41db2c6c",
        "transportWeightWd": 55.0,
        "transportContainerCount": 66.0,
        "turnWeight": 77.0,
        "averageNum": 88.0,
        "cargoType": null,
        "averagePriceHighway": 5.0,
        "averageRoadTransportPrice": 6.0,
        "companyTotalWeightWd": 7.0,
        "companyTotalWeightWCt": 8.0,
        "remark": null,
        "creator": null,
        "createTime": null,
        "modifier": null,
        "modifyTime": null,
        "transportLine": "线路2",
        "transportModeCode": "ROAD_WATER",
        "totalLength": 123.0,
        "roadLength": null,
        "railwayLength": null,
        "waterLength": null
      }]
    }]
          */
    for (let i = 0; i < requestResult.length; i++) {
      const reportMap = requestResult[i]
      let detailList = reportMap.detailList
      if (detailList === undefined || detailList === null) {
        detailList = []
      }
      // 合并，如果子表是0，就不需要合并
      const mergeStep = detailList.length === 0 ? 1 : detailList.length
      // 内容里 需要合并的内容
      mergeContent.push(
        // 序号
        {
          s: {
            r: startMergeLength,
            c: 0
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 0
          }
        },
        // 市别
        {
          s: {
            r: startMergeLength,
            c: 1
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 1
          }
        },
        // 示范项目工程名称
        {
          s: {
            r: startMergeLength,
            c: 2
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 2
          }
        },
        // 线路总数
        {
          s: {
            r: startMergeLength,
            c: 3
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 3
          }
        },
        // 企业同期总货物运输量 万吨
        {
          s: {
            r: startMergeLength,
            c: 20
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 20
          }
        },
        // 企业同期总货物运输量 万标箱
        {
          s: {
            r: startMergeLength,
            c: 21
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 21
          }
        },
        // 枢纽场站建设季度投资完成额（万元）
        {
          s: {
            r: startMergeLength,
            c: 22
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 22
          }
        },
        // 枢纽场站建设累计完成投资额
        {
          s: {
            r: startMergeLength,
            c: 23
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 23
          }
        },
        // 枢纽场站投资完成率
        {
          s: {
            r: startMergeLength,
            c: 24
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 24
          }
        },
        // 装备设备购置季度投资完成额（万元）
        {
          s: {
            r: startMergeLength,
            c: 25
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 25
          }
        },
        // 装备设备购置累计完成投资额
        {
          s: {
            r: startMergeLength,
            c: 26
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 26
          }
        },
        // 装备设备购置完成率
        {
          s: {
            r: startMergeLength,
            c: 27
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 27
          }
        },
        // 信息系统建设季度投资完成额（万元）
        {
          s: {
            r: startMergeLength,
            c: 28
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 28
          }
        },
        // 信息系统建设累计完成投资额
        {
          s: {
            r: startMergeLength,
            c: 29
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 29
          }
        },
        // 信息系统建设完成率
        {
          s: {
            r: startMergeLength,
            c: 30
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 30
          }
        },
        // 其他
        {
          s: {
            r: startMergeLength,
            c: 31
          },
          e: {
            r: startMergeLength + mergeStep - 1,
            c: 31
          }
        }
      )
      startMergeLength += mergeStep
      // 如果没有子表数据，就只展示主表
      if (detailList.length === 0) {
        const formTitle = [
          i + 1, //  序号
          // this.auditStatusByCode(reportMap.auditStatusId), // 审核状态
          reportMap.cityName, // 市别
          reportMap.projectName, // 示范工程项目名称
          detailList.length, // 线路总条数
          1, // 线路
          null, // 联运线路名称
          null, // 联运模式
          null, // 示范路线长度·总长度
          null, // 示范路线长度·公路
          null, // 示范路线长度·铁路
          null, // 示范路线长度·水路
          null, // 联运量·万吨
          null, // 联运量·万标箱
          null, // 联运周转量
          null, // 平均运转次数
          null, // 运输主要货品货类
          null, // 平均单位联运价格
          null, // 平均公路运输价格
          null, // 多式联运量占比%
          null, // 集装箱多式联运量占比%
          null, // 企业同期总货物运输量·万吨
          null, // 企业同期总货物运输量·万标箱
          null, // 联运设施设备·枢纽场站建设季度投资完成额（万元）
          null, // 联运设施设备·枢纽场站建设累计完成投资额
          null, // 联运设施设备·枢纽场站投资完成率
          null, // 联运设施设备·装备设备购置季度投资完成额（万元）
          null, // 联运设施设备·装备设备购置累计完成投资额
          null, // 联运设施设备·装备设备购置完成率
          null, // 联运信息化·信息系统建设季度投资完成额（万元）
          null, // 联运信息化·信息系统建设累计完成投资额
          null, // 联运信息化·信息系统建设完成率
          reportMap.remark // 其他
        ]
        totalForm.push(formTitle)
      } else {
        for (let j = 0; j < detailList.length; j++) {
          const route = detailList[j]
          if (route === undefined || route === null) {
            continue
          }
          // 子表第一个要拼接在主表里
          if (j === 0) {
            const weightWdReduce = detailList.reduce((total, current) => {
              return total + current.transportWeightWd || 0
            }, 0)
            const weightWctReduce = detailList.reduce((total, current) => {
              return total + current.transportContainerCount || 0
            }, 0)

            const formTitle = [
              i + 1, // 序号
              // this.auditStatusByCode(reportMap.auditStatusId), // 审核状态
              reportMap.cityName, // 市别
              reportMap.projectName, // 示范工程项目名称
              detailList.length, // 线路总条数
              j + 1, // 线路
              route.transportLine, // 联运线路名称
              this.target.getModeNameById(route.transportModeCode), // 联运模式
              route.totalLength, // 示范路线长度·总长度
              route.roadLength, // 示范路线长度·公路
              route.railwayLength, // 示范路线长度·铁路
              route.waterLength, // 示范路线长度·水路
              route.transportWeightWd, // 联运量·万吨
              route.transportContainerCount, // 联运量·万标箱
              route.turnWeight, // 联运周转量
              route.averageNum, // 平均运转次数
              route.cargoType, // 运输主要货品货类
              route.averagePriceHighway, // 平均单位联运价格
              route.averageRoadTransportPrice, // 平均公路运输价格
              ((route.transportWeightWd || 0) / (route.companyTotalWeightWd || 1) * 100).toFixed(2), // 多式联运量占比
              ((route.transportContainerCount || 0) / (route.companyTotalWeightWCt || 1) * 100).toFixed(2), // 多式联运量占比
              weightWdReduce, // route.companyTotalWeightWd, // 企业同期总货物运输量·万吨
              weightWctReduce, // route.companyTotalWeightWCt, // 企业同期总货物运输量·万标箱
              reportMap.hubStationBuildInvestMoneyCurrentSeason, // 联运设施设备·枢纽场站建设季度投资完成额（万元）
              reportMap.hubStationBuildInvestMoney, // 联运设施设备·枢纽场站建设累计完成投资额
              reportMap.hubStationBuildInvestRate, // 联运设施设备·枢纽场站投资完成率
              reportMap.equipmentInvestMoneyCurrentSeason, // 联运设施设备·装备设备购置季度投资完成额（万元）
              reportMap.equipmentInvestMoney, // 联运设施设备·装备设备购置累计完成投资额
              reportMap.equipmentInvestRate, // 联运设施设备·装备设备购置完成率
              reportMap.informationSysInvestMoneyCurrentSeason, // 联运信息化·信息系统建设季度投资完成额（万元）
              reportMap.informationSysInvestMoney, // 联运信息化·信息系统建设累计完成投资额
              reportMap.informationSysInvestRate, // 联运信息化·信息系统建设完成率
              reportMap.remark // 其他
            ]
            totalForm.push(formTitle)
          } else {
            const form = [
              null, // 序号
              // null, // 审核状态
              null, // 市别
              null,
              null, // 线路总条数
              j + 1, // 线路
              route.transportLine, // 联运线路名称
              this.target.getModeNameById(route.transportModeCode), // 联运模式
              route.totalLength, // 示范路线长度·总长度
              route.roadLength, // 示范路线长度·公路
              route.railwayLength, // 示范路线长度·铁路
              route.waterLength, // 示范路线长度·水路
              route.transportWeightWd, // 联运量·万吨 是统计量
              route.transportContainerCount, // 联运量·万标箱 是统计量
              route.turnWeight, // 联运周转量
              route.averageNum, // 平均运转次数
              route.cargoType, // 运输主要货品货类
              route.averagePriceHighway, // 平均单位联运价格
              route.averageRoadTransportPrice, // 平均公路运输价格
              ((route.transportWeightWd || 0) / (route.companyTotalWeightWd || 1) * 100).toFixed(2), // 多式联运量占比
              ((route.transportContainerCount || 0) / (route.companyTotalWeightWCt || 1) * 100).toFixed(2), // 多式联运量占比
              null, // route.companyTotalWeightWd, // 企业同期总货物运输量·万吨
              null, // route.companyTotalWeightWCt, // 企业同期总货物运输量·万标箱
              reportMap.hubStationBuildInvestMoney, // 联运设施设备·枢纽场站建设累计完成投资额
              reportMap.hubStationBuildInvestRate, // 联运设施设备·枢纽场站投资完成率
              reportMap.equipmentInvestMoney, // 联运设施设备·装备设备购置累计完成投资额
              reportMap.equipmentInvestRate, // 联运设施设备·装备设备购置完成率
              reportMap.informationSysInvestMoney, // 联运信息化·信息系统建设累计完成投资额
              reportMap.informationSysInvestRate, // 联运信息化·信息系统建设完成率
              null
            ]
            totalForm.push(form)
          }
        }
      }
    }
    // 最后一行的统计
    const itemArray = requestResult.reduce((total, current) => {
      return total.concat(current.detailList)
    }, []).flat()
    const transportWeightWdStatistics = itemArray.reduce((total, current) => {
      return total + (current.transportWeightWd | 0)
    }, 0)
    const transportContainerCountStatistics = itemArray.reduce((total, current) => {
      return total + (current.transportContainerCount | 0)
    }, 0)
    totalForm.push([`          多式联运示范工程项目参与企业总货物运输量（万吨）： ${transportWeightWdStatistics}  ；其中，集装箱货物运输量（万标箱）：${transportContainerCountStatistics}`])
    mergeContent.push(
      // 最后一行的统计
      {
        s: {
          r: startMergeLength,
          c: 0
        },
        e: {
          r: startMergeLength,
          c: 29
        }
      }
    )
    // 合并表头和表内容
    aoa = aoa.concat(totalForm)
    // form => sheet
    const sheet = this.sheet_from_array_of_arrays(aoa)

    // 为某个单元格设置单独样式
    sheet['A2'].s = {
      font: {
        name: '宋体',
        sz: 24,
        bold: true,
        color: { rgb: '000000' }
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrap_text: 'true'
      }
    }
    sheet['A3'].s = {
      font: {
        name: '宋体',
        sz: 13,
        bold: true,
        color: { rgb: '000000' }
      },
      alignment: {
        horizontal: 'left'
      }
    }
    // 最后一行的统计
    sheet[`A${startMergeLength + 1}`].s = {
      font: {
        name: '宋体',
        sz: 13,
        bold: true,
        color: { rgb: '000000' }
      },
      alignment: {
        horizontal: 'left'
      }
    }
    // const style0 = { border: { top: 10, bottom: 10 }, alignment: { horizontal: 'center', wrapText: true, vertical: 'center' }, font: { sz: 18, height: 20, name: '宋体', bold: false, color: { rgb: '000000' }, outline: true }}
    // const style1 = { alignment: { horizontal: 'center', wrapText: true, vertical: 'center' }, font: { sz: 11, name: '宋体', bold: true, color: { rgb: '000000' }, outline: true }}
    // sheet['A1'].s = style0
    // sheet['A2'].s = style1 // 序号
    // sheet['B2'].s = style1 // 市别
    // sheet['C2'].s = style1 // 示范工程项目名称
    // sheet['D2'].s = style1 // 示范工程项目联运线路运营情况
    // sheet['D3'].s = style1 // 线路线路（条）
    // sheet['E3'].s = style1 // 线路
    // sheet['F3'].s = style1 // 联运线路
    // sheet['G3'].s = style1 // 联运模式
    // sheet['H3'].s = style1 // 示范线路长度
    // sheet['H4'].s = style1 // 示范线路长度-总长度（公里）
    // sheet['I4'].s = style1 // 示范线路长度-公里（公里）
    // sheet['J4'].s = style1 // 示范线路长度-铁路（公里）
    // sheet['K4'].s = style1 // 示范线路长度-水路（公里）
    // sheet['L3'].s = style1 // 联运量
    // sheet['L4'].s = style1 // 联运量-万吨
    // sheet['M4'].s = style1 // 联运量-万标箱
    // sheet['N3'].s = style1 // 联运周转量（万吨公里）
    // sheet['O3'].s = style1 // 平均转运次数（次）
    // sheet['P3'].s = style1 // 运输的主要货品货类
    // sheet['Q3'].s = style1 // 平均单位联运价格（元/吨·公里）
    // sheet['R3'].s = style1 // 平均公路运输价格（元/吨·/公里）
    // sheet['S3'].s = style1 // 企业总货物运输量
    // sheet['S4'].s = style1 // 企业同期总货物运输量-万吨
    // sheet['T4'].s = style1 // 企业同期总货物运输量-万标箱
    sheet['!merges'] = mergeTitle.concat(mergeContent)
    // 冻结前6行和第一列，右下可以滑动
    sheet['!freeze'] = {
      xSplit: '1',
      ySplit: '6',
      topLeftCell: 'B7',
      activePane: 'bottomRight',
      state: 'frozen'
    }
    sheet['!margins'] = {
      left: 1.0,
      right: 1.0,
      top: 1.0,
      bottom: 1.0,
      header: 0.5,
      footer: 0.5
    }
    // 列宽 使用的不是像素值
    sheet['!cols'] = [
      { wch: 10 },
      { wch: 14 },
      { wch: 10 },
      { wch: 16 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 }, // 联运模式
      { wch: 16 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 20 }, // 联运周转量（万吨公里）
      { wch: 20 },
      { wch: 20 },
      { wch: 20 },
      { wch: 20 },
      { wch: 14 },
      { wch: 17 },
      { wch: 10 }, // 企业同期总货物运输量	 万吨
      { wch: 10 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 },
      { wch: 38 }
    ]
    // 设置行高，但是没起作用
    /*
          * type RowInfo = {
          hidden?: boolean; // if true, the row is hidden
          /// row height is specified in one of the following ways:
          hpx?:    number;  // height in screen pixels
          hpt?:    number;  // height in points

          level?:  number;  // 0-indexed outline / group level
          };*/
    // sheet['!rows'] = [{ groupLevel: 1, hpx: 300 }, { groupLevel: 3, hpt: 300 }, { groupLevel: 1, hpx: 300 }]
    // sheet['!rows'] = Array(10).fill({ hpx: 100 })
    // sheet => bolb
    const wbBlob = this.sheet2blob(sheet, '线路运营统计表')
    // 保存下载
    FileSaver.saveAs(wbBlob, '江苏省多式联运示范工程项目' + (this.target.form.param.year + '年' || '') + '动态监测信息表.xlsx')
  }

  // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
  sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1'
    const workbook = {
      SheetNames: [sheetName],
      Sheets: {}
    }
    workbook.Sheets[sheetName] = sheet
    // 生成excel的配置项
    const wopts = {
      bookType: 'xlsx', // 要生成的文件类型
      bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
      type: 'binary'
    }

    const wbOut = xlsx.write(workbook, wopts, { defaultCellStyle: this.target.defaultCellStyle })

    // 字符串转ArrayBuffer
    function s2ab(s) {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    }

    return new Blob([s2ab(wbOut)], { type: 'application/octet-stream' })
  }

  // 从json转化为sheet
  sheet_from_array_of_arrays(data) {
    const ws = {}
    const range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    }
    for (let R = 0; R !== data.length; ++R) {
      for (let C = 0; C !== data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R
        if (range.s.c > C) range.s.c = C
        if (range.e.r < R) range.e.r = R
        if (range.e.c < C) range.e.c = C
        const cell = {
          v: data[R][C],
          s: this.target.defaultCellStyle
        }
        if (cell.v == null) continue
        const cell_ref = xlsx.utils.encode_cell({
          c: C,
          r: R
        })

        /* TEST: proper cell types and value handling */
        if (typeof cell.v === 'number') {
          cell.t = 'n'
        } else if (typeof cell.v === 'boolean') {
          cell.t = 'b'
        } else if (cell.v instanceof Date) {
          cell.t = 'n'
          cell.z = xlsx.SSF._table[14]
          cell.v = this.dateNum(cell.v)
        } else {
          cell.t = 's'
        }
        ws[cell_ref] = cell
      }
    }

    /* TEST: proper range */
    if (range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range)
    return ws
  }

  /* TODO: date1904 logic */
  dateNum(v, date1904) {
    if (date1904) v += 1462
    const epoch = Date.parse(v)
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
  }

  validValue(value) {
    return value || '/'
  }
}
export {
  ExportReportExcel
}


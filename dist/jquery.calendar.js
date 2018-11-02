/** 
 * jquery日历插件jqyery.calendar.js 
 * @author:xuzengqiang 
 * @since :2015-2-4 15:31:39 
**/  
;(function($){     
    jQuery.fn.extend({  
        showCalendar:function(options){  
        var defaultOptions = {  
                //class名称  
                className:'strongCalendar',  
                //日历格式'yyyy-MM-dd'('yyyy-MM-dd','yyyy年MM月dd日')  
                format:'yyyy-MM-dd',  
                //高度,默认220px  
                height:220,  
                //宽度:默认与文本框宽度相同  
                width:$(this).innerWidth(),  
                //日历框离文本框高度  
                marginTop:1,  
                //浮层z-index  
                zIndex:99,  
                //间隙距离,默认为5px  
                spaceWidth:8,  
                //字体大小,默认9pt  
                fontSize:9,  
                //日历背景色  
                bgColor:'#FFFFFF',  
                //默认浅灰色  
                borderColor:"#AFAFAF",  
                //顶部背景颜色,默认为淡灰色  
                topColor:'#FFFFFF',  
                //当前年月字体颜色  
                ymFontColor:'#535353',  
                //年月份操作背景色  
                ymBgColor:'#FFFFFF',  
                //年月份移上背景色  
                ymHoverBgColor:'#EFEFEF',  
                //箭头颜色  
                arrowColor:'#535353',  
                //里层边框  
                innerBorder:'1px solid #AFAFAF',  
                //表格边框  
                tableBorder:'0px solid #AFAFAF',  
                //星期背景颜色  
                weekBgColor:'#EFEFEF',  
                //星期字体颜色  
                weekFontColor:'#535353',  
                //上个月和下个月日期的字体颜色  
                noThisMonthFontColor:'#CFCFCF',  
                //这个月的日期字体颜色  
                thisMonthFontColor:'#535353',  
                //这个的日期背景颜色  
                thisMonthBgColor:'#FFFFFF',  
                //日期移上时字体颜色  
                dateHoverFontColor:'#3399CC',  
                //日期移上时字体颜色  
                dateHoverBgColor:'#EFEFEF',  
                //button边框  
                btnBorder:'1px solid #AFAFAF',  
                //button背景色  
                btnBgColor:'#FFFFFF',  
                //button字体颜色  
                btnFontColor:'#535353',  
                //button鼠标移上颜色  
                btnHoverBgColor:'#EFEFEF',  
                //button鼠标移上字体颜色  
                btnHoverFontColor:'#3399CC'  
            };  
            var settings = jQuery.extend(defaultOptions,options || {}),  
                $body = $("body"),  
                date = new Date(),  
                currentYear = date.getFullYear(),  
                currentMonth = date.getMonth(),  
                monthDay = [],  
                //行高  
                lineHeight = parseInt(settings.height-4*settings.spaceWidth)/9,  
                $calendar,  
                //触发事件对象  
                $target = $(this),  
                current ;  
              
            //对于小于10的数字前+0  
            Number.prototype.addZero = function(){  
                return this<10?"0"+this:this;  
            };  
            var Calendar = {  
                    //星期列表  
                    weeks : ['日','一','二','三','四','五','六'],  
                    //每月的天数  
                    months : [31,0,31,30,31,30,31,31,30,31,30,31],  
                    //初始化日历  
                    loadCalendar:function(){  
                        $body.append("<div class="+settings.className+"></div>");  
                        $calendar = $("."+settings.className);  
                        //新增核心html  
                        $calendar.append(Calendar.innerHTML());  
                        //样式加载  
                        Calendar.styleLoader();  
                        //核心日历加载  
                        Calendar.loaderDate(currentYear,currentMonth);  
                        //事件绑定  
                        Calendar.dateEvent();  
                    },  
                    //日历核心HTML  
                    innerHTML:function(){  
                        var htmlContent = {};  
                        htmlContent = "<div class='cal_head'>"+ <!--头部div层start-->  
                                          "<div class='year_oper oper_date'>"+  
                                            "<div class='year_sub operb'><b class='arrow aLeft'/></div>"+  
                                            "<div class='year_add operb'><b class='arrow aRight'/></div>"+  
                                            "<span class='current_year'></span>"+   
                                            <!--位置互换下,考虑到span后面float:right会换行-->  
                                          "</div>"+  
                                          "<div class='month_oper oper_date'>"+  
                                            "<div class='month_sub operb'><b class='arrow aLeft'/></div>"+  
                                            "<div class='month_add operb'><b class='arrow aRight'/></div>"+  
                                            "<span class='current_month'></span>"+  
                                          "</div>"+  
                                      "</div>"+ <!--头部div层end-->  
                                      "<div class='cal_center'><table cellspacing='0'></table></div>"+  
                                      "<div class='cal_bottom'>"+ <!--底部div层start-->  
                                        "<button class='clear_date'>清空</button>"+  
                                        "<button class='today_date'>今天</button>"+  
                                        "<button class='confirm_date'>确定</button>"+  
                                      "</div>";<!--底部div层end-->  
                        return htmlContent;  
                                     
                    },  
                    //日历样式加载  
                    styleLoader:function(){  
                        $calendar = $("."+settings.className);    
                        //总弹出层样式设置  
                        $calendar.css({"border-width":"1px",  
                                       "border-color":settings.borderColor,  
                                       "background-color":settings.bgColor,  
                                       "border-style":"solid",  
                                       "height":settings.height,  
                                       "width":settings.width,  
                                       "z-index":settings.zIndex,  
                                       "font-size":settings.fontSize+"pt"  
                                      });  
                        Calendar.setLocation();  
                        //日历顶部样式设置  
                        var $calHead=$calendar.find(".cal_head"),  
                            $operDate=$calendar.find(".oper_date"),  
                            $arrow=$calHead.find(".arrow"),  
                            $center=$calendar.find(".cal_center"),  
                            $ctable=$center.find("table"),  
                            //箭头大小  
                            arrowWidth = 6,  
                            $calBottom = $calendar.find(".cal_bottom");  
                              
                        $calHead.css({"height":lineHeight+2*settings.spaceWidth,  
                                      "background-color":settings.topColor  
                                     });  
                          
                        $operDate.css({"margin-top":settings.spaceWidth,  
                                       "margin-left":settings.spaceWidth,  
                                       "float":"left",  
                                       "border":settings.innerBorder,  
                                       "text-align":"center"  
                                     });  
                        //设置operDate包含边框的实际宽度  
                        $operDate.outerWidth(($calHead.width()-3*settings.spaceWidth)/2);  
                        $operDate.outerHeight(lineHeight);  
                        $operDate.find(".operb")  
                                 .css({"width":"20px",  
                                       "background":settings.ymBgColor,  
                                       "height":$operDate.innerHeight(),  
                                       "cursor":"pointer"  
                                     });  
                        $operDate.find(".year_sub,.month_sub")  
                                 .css({"float":"left","border-right":settings.innerBorder});  
                        $operDate.find(".year_add,.month_add")  
                                 .css({"float":"right","border-left":settings.innerBorder});  
                        $operDate.find("span")  
                                 .css({"color":settings.ymFontColor,  
                                       "line-height":$operDate.height()+"px"  
                                     });  
                                   
                        //设置箭头样式  
                        $operDate.find(".aLeft")  
                                 .arrow({"direction":"left",  
                                         "width":arrowWidth,  
                                         "height":arrowWidth*2,  
                                         "color":settings.arrowColor  
                                       });  
                        $operDate.find(".aRight")  
                                 .arrow({"direction":"right",  
                                         "width":arrowWidth,  
                                         "height":arrowWidth*2,  
                                         "color":settings.arrowColor  
                                       });  
                        $arrow.css({  
                            "margin":"0 auto",  
                            "margin-top":parseInt($operDate.innerHeight())/2-arrowWidth     
                        });  
                          
                        //加载中间表格样式  
                        $center.css({  
                            "border":settings.innerBorder,  
                            "margin-left":settings.spaceWidth,  
                            "overflow":'hidden'  
                        });  
                        //加载中间表格实际宽度  
                        $center.outerWidth($calendar.width()-2*settings.spaceWidth);  
                        $center.height(lineHeight*7);  
                          
                        $ctable.find("td").css({"text-align":"center"});  
                        $calBottom.css({"margin-right":settings.spaceWidth});  
                        $calBottom.find("button")  
                                  .css({"border":settings.btnBorder,  
                                        "background":settings.btnBgColor,  
                                        "color":settings.btnFontColor,  
                                        "margin-top":settings.spaceWidth,  
                                        "margin-left":settings.spaceWidth,  
                                        "float":"right","width":"20%"  
                                      });  
                        $calBottom.find("button").outerHeight(lineHeight);  
                        //去除button链接的虚线框  
                        $("."+settings.className+" button").focus(function(){this.blur()});  
                          
                        $operDate.find(".operb").hover(function(){  
                            $(this).css("background",settings.ymHoverBgColor);                         
                        },function(){  
                            $(this).css("background",settings.ymBgColor);  
                        });  
                          
                        $calBottom.find("button").hover(function(){  
                            $(this).css({"background":settings.btnHoverBgColor,"color":settings.btnHoverFontColor});  
                        },function(){  
                            $(this).css({"background":settings.btnBgColor,"color":settings.btnFontColor});  
                        });  
                    },  
                    //事件绑定  
                    dateEvent:function(){  
                        var $calendar = $("."+settings.className);  
                        $calendar.find(".year_add").click(function(){Calendar.yearAdd();});  
                        $calendar.find(".year_sub").click(function(){Calendar.yearSub();});  
                        $calendar.find(".month_add").click(function(){Calendar.monthAdd();});  
                        $calendar.find(".month_sub").click(function(){Calendar.monthSub();});  
                        $calendar.find(".confirm_date").click(function(){Calendar.confirm();});  
                        $calendar.find(".today_date").click(function(){Calendar.getToday();});  
                        $calendar.find(".clear_date").click(function(){Calendar.clear();});  
                    },  
                    //当前日期:年份和月份  
                    date:function(){  
                        var $calendar = $("."+settings.className);  
                        return {  
                            year:parseInt($calendar.find(".current_year").html()),  
                            month:parseInt($calendar.find(".current_month").html())  
                        };  
                    },  
                    //年份自增  
                    yearAdd:function(){  
                        Calendar.loaderDate(Calendar.date().year+1,Calendar.date().month-1);  
                    },  
                    //年份自减  
                    yearSub:function(){  
                        Calendar.loaderDate(Calendar.date().year-1,Calendar.date().month-1);  
                    },  
                    //月份自增  
                    monthAdd:function(){  
                        var year = Calendar.date().year, month = Calendar.date().month;  
                        if(month==12) {    
                            month=1;    
                            year=year+1;    
                        } else {    
                            month=month+1;    
                        }    
                        Calendar.loaderDate(year,month-1);  
                    },  
                    //月份自减  
                    monthSub:function(){  
                        var year = Calendar.date().year, month = Calendar.date().month;  
                        if(month==1) {    
                            month=12;    
                            year=year-1;    
                        } else {    
                            month=month-1;    
                        }    
                        Calendar.loaderDate(year,month-1);  
                    },  
                    //日期选择  
                    dateChoose:function($object){  
                        var year = Calendar.date().year, month = Calendar.date().month;  
                        //上一个月  
                        if($object.hasClass("pre_month_day")) {  
                            if(month == 1) {  
                                year = year-1;  
                                month = 12;  
                            } else {  
                                month = (month-1).addZero();  
                            }  
                        //当前月     
                        } else if($object.hasClass("this_month_day")) {  
                            month = month.addZero();  
                        //下一个月  
                        } else {  
                            if(month == 12) {  
                                month = "01";  
                                year = year + 1;  
                            } else {  
                                month = (month+1).addZero();  
                            }  
                        }  
                        current.val(year+"-"+month+"-"+$object.text());  
                    },  
                    //确定事件  
                    confirm:function(){  
                        Calendar.destory();  
                    },  
                    //是否是闰年  
                    isLeapYear:function(year){  
                        if((year%4==0 && year%100!=0) || (year%400==0)) {  
                            return true;      
                        }  
                        return false;  
                    },      
                    //指定年份二月的天数  
                    februaryDays:function(year){  
                        if(typeof year !== "undefined" && parseInt(year) === year) {  
                            return Calendar.isLeapYear(year) ? 29:28;  
                        }  
                        return false;  
                    },  
                    getWeek:function(num){  
                        return Calendar.weeks[num];  
                    },  
                    //获取指定月份[0,11]的天数  
                    getMonthDay:function(year,month){  
                        if(month === 1){  
                            return Calendar.februaryDays(year);  
                        }  
                        month=(month===-1)?11:month;  
                        return Calendar.months[month];  
                    },  
                    //今天  
                    getToday:function(){  
                        var date = new Date(),  
                            year = date.getFullYear(),  
                            month = (date.getMonth()+1).addZero(),  
                            day = date.getDate().addZero();  
                        current.val(year+"-"+month+"-"+day);  
                        Calendar.destory();  
                    },  
                    //清空  
                    clear:function(){  
                        current.val("");  
                        Calendar.destory();  
                    },  
                    //设置日历位置  
                    setLocation:function(){  
                        $calendar = $("."+settings.className);    
                        var left = current.offset().left,  
                            top = current.offset().top + current.innerHeight() + settings.marginTop;  
                        $calendar.css({"position":"absolute","left":left,"top":top});  
                    },  
                    //关闭日历  
                    destory:function(){  
                        $("."+settings.className).empty().remove();  
                    },  
                    //初始化日历主体内容  
                    loaderDate:function(year,month){  
                        //初始化日期为当前年当前月的1号,获取1号对应的星期     
                        var oneWeek=new Date(year,month,1).getDay(),  
                            $calendar = $("."+settings.className),  
                            $calTable = $calendar.find("table"),  
                            //这个月天数  
                            thisMonthDay = Calendar.getMonthDay(year,month),  
                            //获取上一个的天数  
                            preMonthDay;  
                        //清空日期列表  
                        $calTable.html("");  
                          
                        $calendar.find(".current_year").text(year+"年");   
                        $calendar.find(".current_month").text((month+1)+"月");     
                        if(oneWeek == 0) oneWeek = 7;  
                          
                        if(i==0) {  
                            preMonthDay = Calendar.getMonthDay(year-1,11);  
                        } else {  
                            preMonthDay = Calendar.getMonthDay(year,month-1);  
                        }  
                        var start = 1, end = 1;    
                        for(var i=0;i<7;i++) {  
                            var dayHTML = "";  
                            if(i==0) {  
                                $calTable.append("<tr class='week_head'><tr>");  
                            }  
                            for(var j=1;j<=7;j++) {    
                                //设置星期列表  
                                if(i==0) {     
                                    $calTable.find(".week_head").append("<td>"+Calendar.getWeek(j-1)+"</td>");   
                                } else {    
                                    if((i-1)*7+j<=oneWeek) { //从第二行开始设置值                         
                                        dayHTML+="<td class='pre_month_day'>"+(preMonthDay-oneWeek+j)+"</td>";  
                                    } else if((i-1)*7+j<=thisMonthDay+oneWeek ){    
                                        var result=(start++).addZero();    
                                        dayHTML+="<td class='this_month_day'>"+result+"</td>";  
                                    } else {    
                                        var result=(end++).addZero();    
                                        dayHTML +="<td class='next_month_day'>"+result+"</td>";   
                                    }    
                                }     
                            }       
                            if(i>0){  
                                $calTable.append("<tr>"+dayHTML+"</tr>");  
                            }  
                        }   
                        Calendar.tableStyle();  
                    },  
                    //表格样式初始化  
                    tableStyle:function(){  
                        //表格中单元格的宽度  
                        var $center = $calendar.find(".cal_center"),  
                            $calTable = $calendar.find("table"),  
                            tdWidth = parseFloat($center.width())/7;  
                              
                        $calTable.find("td").css({"width":tdWidth,"text-align":"center",  
                                                  "color":"#AFAFAF",  
                                                  "background":settings.thisMonthBgColor,"cursor":"pointer",  
                                                  "color":settings.thisMonthFontColor,  
                                                  "border-top":settings.tableBorder,  
                                                  "border-right":settings.tableBorder});  
                        $calTable.find(".week_head td")  
                                 .css({"background":settings.weekBgColor,  
                                       "cursor":"auto","border":"0",  
                                       "color":settings.weekFontColor  
                                     });  
                        $calTable.find(".pre_month_day,.next_month_day")  
                                 .css({"color":settings.noThisMonthFontColor,"background":"transparent"});  
                        //去除最右侧表格的右边框  
                        $calTable.find("td:nth-child(7n)").css({"border-right":0});  
                        $calTable.find("td").outerHeight(lineHeight);  
                      
                        $calTable.find("tr[class!=week_head] td").hover(function(){  
                            $(this).css({"background":settings.dateHoverBgColor,  
                                         "color":settings.dateHoverFontColor  
                                       });  
                        },function(){  
                            $(this).css({"background":settings.thisMonthBgColor});  
                            if($(this).hasClass("this_month_day")) {  
                                $(this).css({"color":settings.thisMonthFontColor});  
                            } else {  
                                $(this).css({"color":settings.noThisMonthFontColor});  
                            }  
                        }).click(function(){Calendar.dateChoose($(this));});  
                    }  
                };  
              
            return this.each(function(){  
                $target.click(function(){  
                    //触发对象,请写在点击事件中,不然会做为全局变量!  
                    current = $(this);  
                    //如果没有生成日历元素  
                    if($("."+settings.className).length == 0) {  
                        Calendar.loadCalendar();                     
                    }  
                });  
                //事件触发对象  
                $(document).click(function(event){  
                    var $calendar = $("."+settings.className)  
                    if(!$target.triggerTarget(event) && !$calendar.triggerTarget(event)) {  
                        Calendar.destory();  
                    }  
                });  
            });  
        },  
        /** 
         *描述:生成指定箭头样式 
        **/  
        arrow:function(options){  
            var defaultOptions = {  
                    color:'#AFAFAF',  
                    height:20,  
                    width:20,  
                    //方向:默认向上'top',供选择['up','bottom','left','right']  
                    direction:'top'  
                };  
            var settings = jQuery.extend(defaultOptions,options||{}),  
                current = $(this);  
            function loadStyle(){  
                current.css({"display":"block","width":"0","height":"0"});  
                if(settings.direction === "top" || settings.direction === "down") {  
                    current.css({  
                                "border-left-width":settings.width/2,  
                                "border-right-width":settings.width/2,  
                                "border-left-style":"solid",  
                                "border-right-style":"solid",  
                                "border-left-color":"transparent",  
                                "border-right-color":"transparent"  
                                });  
                    if(settings.direction === "top") {  
                        current.css({  
                                "border-bottom-width":settings.height,  
                                "border-bottom-style":"solid",  
                                "border-bottom-color":settings.color      
                                });  
                    } else {  
                        current.css({  
                                "border-top-width":settings.height,  
                                "border-top-style":"solid",  
                                "border-top-color":settings.color     
                                });  
                    }  
                } else if(settings.direction === "left" || settings.direction === "right") {  
                    current.css({  
                                "border-top-width":settings.height/2,  
                                "border-bottom-width":settings.height/2,  
                                "border-top-style":"solid",  
                                "border-bottom-style":"solid",  
                                "border-top-color":"transparent",  
                                "border-bottom-color":"transparent"  
                                });  
                    if(settings.direction === "left") {  
                        current.css({  
                                "border-right-width":settings.width,  
                                "border-right-style":"solid",  
                                "border-right-color":settings.color   
                                });  
                    } else {  
                        current.css({  
                                "border-left-width":settings.width,  
                                "border-left-style":"solid",  
                                "border-left-color":settings.color    
                                });  
                    }  
                }  
            }  
            return this.each(function(){ loadStyle(); });  
        },  
        //触发事件是否是对象本身  
        triggerTarget:function(event){  
            //如果触发的是事件本身或者对象内的元素  
            return $(this).is(event.target) || $(this).has(event.target).length > 0;   
        },  
        //皮肤  
        skin:function(){  
              
        }  
    });  
})(jQuery);  
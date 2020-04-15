(function ($) {
    let cols, rows, td_width, td_height, that, setting, copy_board;
    // 注册JQ函数
    $.fn.extend({
        excel: function (options) {
            that = this;
            setting = $.extend({}, options);
            // 设置默认列宽
            td_width = setting.cellWidth ? setting.cellWidth : 80;
            // 设置默认行高
            td_height = setting .cellHeight ? setting.cellHeight : 30;

            // 绑定事件
            initBind();
            // 初始化控件
            initExcel();
        }
    });

    // 初始化控件
    function initExcel()
    {
        that.resize();
        that.addClass('excel');
    }

    // 表格内容操作事件
    function excelMainTableBody()
    {
        // 获取元素名
        let getName = function(child = ''){
            return '.excel-main-table-data.active .excel-main-table-body' + child;
        };
        // 获取当前元素
        let present = function(child = ''){
            return that.find('.excel-main-table-data.active .excel-main-table-body' + child);
        };
        
        // 选取框移动事件
        that.off('mousedown', getName(' td')).on('mousedown', getName(' td'), function(e){
            if(e.which == 1)
            {
                let t_y = $(this).position().top;
                let t_x = $(this).position().left;

                // 设置选中元素X坐标
                let x1 = $(this).index() + 1;
                // 设置选中元素Y坐标
                let y1 = $(this).parent().index() + 1;

                // 获取选取框
                let cheek = present().find('.excel-main-table-cheek');
                if(cheek.length == 0)
                {
                    present().append('<div class="excel-main-table-cheek"></div>');
                    cheek = present().find('.excel-main-table-cheek');
                }
                // 设置选取框选中当前元素
                cheek.attr({'x1':x1,'y1':y1,'x2':x1,'y2':y1}).css({'left': t_x, 'top': t_y, 'width':getCheekWidth(x1,x1), 'height':getCheekHeight(y1, y1)});
                that.find('input[name="excel-head-input"]').val($(this).html());
                that.find('.excel-main-table-data.active .excel-main-table-select').html(getColumn(x1)+y1);

                $(document).mousemove(function(event) {
                    that.off('mouseover', getName(' td')).on('mouseover', getName(' td'), function(e){
                        let x2 = $(this).index() + 1;
                        let y2 = $(this).parent().index() + 1; 
                        cheek.css('width',getCheekWidth(x1, x2));
                        cheek.css('height',getCheekHeight(y1, y2));
                        if(x2 >= x1)
                        {
                            cheek.css('left',t_x).attr({'x1':x1,'x2':x2});
                        }else{
                            cheek.css('left',t_x - getCheekWidth(x2, x1-1)).attr({'x1':x2,'x2':x1});
                        }
                        if(y2 >= y1)
                        {
                            cheek.css('top',t_y).attr({'y1':y1,'y2':y2});
                        }else{
                            cheek.css('top',t_y - getCheekHeight(y2, y1-1)).attr({'y1':y2,'y2':y1});
                        }

                        if(x2 == x1)
                        {
                            that.find('.excel-main-table-data.active .excel-main-table-select').html(getColumn(x1)+y1);
                        }else if(x2 > x1)
                        {
                            that.find('.excel-main-table-data.active .excel-main-table-select').html(getColumn(x1)+y1 + '*'+getColumn(x2)+y2);
                        }else{
                            that.find('.excel-main-table-data.active .excel-main-table-select').html(getColumn(x2)+y2+ '*'+ getColumn(x1)+y1);
                        }
                        that.find('input[name="excel-head-input"]').val($(this).html());
                    });
                });
                $(document).mouseup(function(){
                    $(document).off('mousemove');
                    that.off('mouseover', getName(' td'));
                });
            }
        });

        function getCheekHeight(y1, y2)
        {
            let height = 0;
            let temp = y1;

            if(y1 > y2)
            {
                y1 = y2;
                y2 = temp;
            }
            
            let temp_tr = present().find('tr');
            if(y1 > 0 && y2 > 0 && y1 <= temp_tr.length && y2 <= temp_tr.length)
            {
                for(let i=y1-1; i<y2; i++)
                {
                    height += temp_tr.eq(i).outerHeight();
                }
            }
            return height;

        }

        function getCheekWidth(x1, x2)
        {
            let width = 0;
            let temp = x1;

            if(x1 > x2)
            {
                x1 = x2;
                x2 = temp;
            }
            
            let temp_td = present().find('tr:first td');
            if(x1 > 0 && x2 > 0 && x1 <= temp_td.length && x2 <= temp_td.length)
            {
                for(let i=x1-1; i<x2; i++)
                {
                    width += temp_td.eq(i).outerWidth();
                }
            }
            return width;
        }

        that.off('click', getName(' td')).on('click', getName(' td'), function(){
            let t_y = $(this).position().top;
            let t_x = $(this).position().left;

            // 设置选中元素X坐标
            let x1 = $(this).index() + 1;
            // 设置选中元素Y坐标
            let y1 = $(this).parent().index() + 1;

            // 获取选取框
            let cheek = present().find('.excel-main-table-cheek');
            if(cheek.length == 0)
            {
                present().append('<div class="excel-main-table-cheek"></div>');
                cheek = present().find('.excel-main-table-cheek');
            }
            // 设置选取框选中当前元素
            cheek.attr({'x1':x1,'y1':y1,'x2':x1,'y2':y1}).css({'left': t_x, 'top': t_y, 'width':getCheekWidth(x1,x1), 'height':getCheekHeight(y1, y1)});
            that.find('input[name="excel-head-input"]').val($(this).html());
            that.find('.excel-main-table-data.active .excel-main-table-select').html(getColumn(x1)+y1);
        });

        that.off('dblclick', getName(' td')).on('dblclick', getName(' td'), function(){
            let x = $(this).index() + 1;
            let y = $(this).parent().index() + 1;
            $(this).attr('contentEditable', true);
            $(this).focus();
            if (document.body.createTextRange) {
                var range = document.body.createTextRange();
                range.moveToElementText(this);
                range.select();
            } else if (window.getSelection) {
                var selection = window.getSelection();
                var range = document.createRange();
                range.selectNodeContents(this);
                selection.removeAllRanges();
                selection.addRange(range);
            }
            $(this).blur(function(){
                $(this).attr('contentEditable', false);
            });
        });

        // 右键菜单事件
        that.off('contextmenu', getName()).on('contextmenu', getName(), function(e) {
            that.append('<div class="contextmenu"><div data="clipboard"><span class="icon">&#xe600;</span>剪切单元格</div><div data="copy"><span class="icon">&#xe621;</span>复制单元格</div><div data="paste" class="hr"><span class="icon">&#xe601;</span>粘贴单元格</div><div data="iUpper"><span class="icon">&#xe653;</span>上方插入一行</div><div data="iDown"><span class="icon">&#xe65a;</span>下方插入一行</div><div data="iLeft"><span class="icon">&#xe65d;</span>左边插入一列</div><div data="iRight" class="hr"><span class="icon">&#xe65b;</span>右边插入一列</div><div data="dRow"><span class="icon">&#xe6d2;</span>删除行</div><div data="dCol"><span class="icon">&#xe6d3;</span>删除列</div></div>');
			// 获取窗口尺寸
			var winWidth = $(document).width();
			var winHeight = $(document).height();
			// 鼠标点击位置坐标
			var mouseX = e.pageX;
			var mouseY = e.pageY;
			// ul标签的宽高
			var menuWidth = $(".contextmenu").width();
			var menuHeight = $(".contextmenu").height();
			// 最小边缘margin(具体窗口边缘最小的距离)
			var minEdgeMargin = 10;
			// 以下判断用于检测ul标签出现的地方是否超出窗口范围
			// 第一种情况：右下角超出窗口
			if(mouseX + menuWidth + minEdgeMargin >= winWidth &&
				mouseY + menuHeight + minEdgeMargin >= winHeight) {
				menuLeft = mouseX - menuWidth - minEdgeMargin + "px";
				menuTop = mouseY - menuHeight - minEdgeMargin + "px";
			}
			// 第二种情况：右边超出窗口
			else if(mouseX + menuWidth + minEdgeMargin >= winWidth) {
				menuLeft = mouseX - menuWidth - minEdgeMargin + "px";
				menuTop = mouseY + minEdgeMargin + "px";
			}
			// 第三种情况：下边超出窗口
			else if(mouseY + menuHeight + minEdgeMargin >= winHeight) {
				menuLeft = mouseX + minEdgeMargin + "px";
				menuTop = mouseY - menuHeight - minEdgeMargin + "px";
			}
			// 其他情况：未超出窗口
			else {
				menuLeft = mouseX + minEdgeMargin + "px";
				menuTop = mouseY + minEdgeMargin + "px";
			};
            // ul菜单出现
            that.find('.contextmenu').css({
				"left": menuLeft,
				"top": menuTop
            }).show();
            // 点击之后，右键菜单隐藏
            that.off('click', '.contextmenu div').on('click', '.contextmenu div', function(){
                let mold = $(this).attr('data');
                if(mold == 'clipboard')
                {
                    document.addEventListener('cut', cut);
                    document.execCommand('cut');
                    document.removeEventListener('cut', cut);
                }else if(mold == 'copy')
                {
                    document.addEventListener('copy', copy);
                    document.execCommand('copy');
                    document.removeEventListener('copy', copy);
                }else if(mold == 'paste')
                {
                    excelPaste(copy_board);
                }else if($.inArray(mold, ['iUpper', 'iDown', 'iLeft','iRight']) > -1)
                {
                    let cheek = present().find('.excel-main-table-cheek');
                    let x = parseInt(cheek.attr('x1')) - 1;
                    let y = parseInt(cheek.attr('y1')) - 1;
                    insertTable(x, y, mold);
                }
                

            });
            $(document).click(function() {
                $(".contextmenu").remove();
            });
			// 阻止浏览器默认的右键菜单事件
			return false;
        });

        function insertTable(x, y, mold)
        {
            // 获取当前元素
            let present_active = function(child = ''){
                return that.find('.excel-main-table-data.active' + child);
            };
            let temp_tr = present(' tr');
            let left_tr = present_active(' .excel-main-table-left tr');
            let head_table = present_active(' .excel-main-table-head table');
            let head_td = present_active(' .excel-main-table-head tr').first().find('td');
            
            // 上方插入
            if(mold == 'iUpper')
            {
                left_tr.last().after('<tr style="height: '+td_height+'px;"><td>'+(left_tr.length+1)+'</td></tr>');
                temp_tr.eq(y).before('<tr style="height: '+td_height+'px;">' + '<td></td>'.repeat(present(' tr:first td').length) + '</tr>');
            }else if(mold == 'iDown')
            {
                left_tr.last().after('<tr style="height: '+td_height+'px;"><td>'+(left_tr.length+1)+'</td></tr>');
                temp_tr.eq(y).after('<tr style="height: '+td_height+'px;">' + '<td></td>'.repeat(present(' tr:first td').length) + '</tr>');
            }else if(mold == 'iLeft')
            {
                temp_tr.eq(0).find('td').eq(x).before('<td style="width: '+td_width+'px;"></td>');
                for(let i = 1; i < temp_tr.length; i++)
                {
                    temp_tr.eq(i).find('td').eq(x).before('<td></td>');
                }
                head_td.eq(x).before('<td style="width: '+td_width+'px;">'+getColumn(x+1)+'</td>');
                for(let i = x; i <= head_td.length; i++)
                {
                    head_td.eq(i).html(getColumn(i+2));
                }
                head_table.width(head_table.width());
                present(' table').width(present(' table').width());
            }else if(mold == 'iRight')
            {
                temp_tr.eq(0).find('td').eq(x).after('<td style="width: '+td_width+'px;"></td>');
                for(let i = 1; i < temp_tr.length; i++)
                {
                    temp_tr.eq(i).find('td').eq(x).after('<td></td>');
                }
                head_td.eq(x).after('<td style="width: '+td_width+'px;">'+getColumn(x+2)+'</td>');
                for(let i = x+1; i <= head_td.length; i++)
                {
                    head_td.eq(i).html(getColumn(i+2));
                }
                head_table.width(head_table.width());
                present(' table').width(present(' table').width());
            }
            temp_tr = present(' tr').eq(y).find('td').eq(x).click();
            that.resize();
        }

        let temp_head_td = '.excel-main-table-data.active .excel-main-table-head td';
        that.off('mouseover', temp_head_td).on('mouseover', temp_head_td, function(e) {
            let temp_this = $(this);
            $(document).mousemove(function(event) {
                that.off('mousedown', temp_head_td).on('mousedown', temp_head_td, function(e){
                    let start_pageX = e.pageX;
                    if(e.which == 1)
                    {
                        $(document).mousemove(function(event) {
                            let temp_td_width = temp_this.width() + (event.pageX - start_pageX);
                            temp_this.width(temp_td_width);
                            present(' tr:first td').eq(temp_this.index()).width(temp_td_width);
                            let cheek = present().find('.excel-main-table-cheek');
                            if(cheek.length > 0)
                            {
                                let x = parseInt(cheek.attr('x1')) - 1;
                                let y = parseInt(cheek.attr('y1')) - 1;
                                present(' tr').eq(y).find('td').eq(x).click();
                            }
                            start_pageX = event.pageX;
                        });
                        $(document).mouseup(function(){
                            $(document).off('mousemove');
                        });
                    }
                });
                
            });
        });
        that.off('mouseover', getName()).on('mouseover', getName(), function(e) {
            document.addEventListener('cut', cut);
            document.addEventListener('copy', copy);
            document.addEventListener('paste', paste);
        }).off('mouseout', getName()).on('mouseout', getName(), function(e) {
            document.removeEventListener('cut', cut);
            document.removeEventListener('copy', copy);
            document.removeEventListener('paste', paste);
        });

        function cut(event)
        {
            if (!(event.clipboardData && event.clipboardData.items))
            {
                return;
            }
            let cheek = present().find('.excel-main-table-cheek');
            if(cheek.length > 0)
            {
                let x = parseInt(cheek.attr('x1')) - 1;
                let y = parseInt(cheek.attr('y1')) - 1;
                let n = parseInt(cheek.attr('x2')) - x;
                let m = parseInt(cheek.attr('y2')) - y;

                let text = [];
                for(let i=0; i<m; i++)
                {
                    let temp_text = [];
                    for(let j=0; j<n; j++)
                    {
                        let temp_this = present(' tr').eq(y+i).find('td').eq(x+j);
                        if(temp_this.length > 0)
                        {
                            temp_text.push(temp_this.html());
                            temp_this.html('');
                        }
                    }
                    text.push(temp_text.join('\t'));
                }
                copy_board = text.join('\n');
                event.clipboardData.setData('text/plain', text.join('\n'));
                event.preventDefault();
            }
        }
        function copy(event)
        {
            if (!(event.clipboardData && event.clipboardData.items))
            {
                return;
            }
            let cheek = present().find('.excel-main-table-cheek');
            if(cheek.length > 0)
            {
                let x = parseInt(cheek.attr('x1')) - 1;
                let y = parseInt(cheek.attr('y1')) - 1;
                let n = parseInt(cheek.attr('x2')) - x;
                let m = parseInt(cheek.attr('y2')) - y;

                let text = [];
                for(let i=0; i<m; i++)
                {
                    let temp_text = [];
                    for(let j=0; j<n; j++)
                    {
                        let temp_this = present(' tr').eq(y+i).find('td').eq(x+j);
                        if(temp_this.length > 0)
                        {
                            temp_text.push(temp_this.html());
                        }
                    }
                    text.push(temp_text.join('\t'));
                }
                copy_board = text.join('\n');
                event.clipboardData.setData('text/plain', text.join('\n'));
                event.preventDefault();
            }
        }
        function paste(event){
            if (!(event.clipboardData && event.clipboardData.items))
            {
                return;
            }
            if(event.clipboardData.items.length > 0)
            {
                var item = event.clipboardData.items[0];
                if (item.kind === "string")
                {
                    item.getAsString(function(str){
                        excelPaste(str);
                    });
                }
            }
        }

        function excelPaste(str)
        {
            if(typeof str === "string")
            {
                str = str.split('\n');
                let cheek = present().find('.excel-main-table-cheek');
                if(cheek.length > 0)
                {
                    let x = parseInt(cheek.attr('x1')) - 1;
                    let y = parseInt(cheek.attr('y1')) - 1;
                    let temp_tr = present(' tr');
                    if(y+str.length > temp_tr.length)
                    {
                        for(let i=temp_tr.length; i < y+str.length - 1; i++)
                        {
                            insertTable(0, i-1, 'iDown');
                        }
                    }

                    for(let i=0; i<str.length; i++)
                    {
                        let temp = str[i].split('\t');
                        if(temp.length != 1 && temp[0] != "")
                        {
                            let temp_td = present(' tr').eq(y+i).find('td');

                            if(x+temp.length > temp_td.length)
                            {
                                for(let j= temp_td.length; j< x+temp.length; j++)
                                {
                                    insertTable(j-1, 0, 'iRight');
                                }
                            }

                            for(let j=0; j<temp.length; j++)
                            {
                                let temp_this = present(' tr').eq(y+i).find('td').eq(x+j);
                                if(temp_this.length > 0)
                                {
                                    temp_this.html(temp[j].replace('\r', ''));
                                }
                            }
                        }
                    }
                }
            }
        }

        that.on("mousewheel DOMMouseScroll",'.excel-main-table-data.active .excel-main-table-body-and-left', function (event){
            let temp, top, move, cha;
            let sheet_active = that.find('.excel-main-table-data.active');
            let body_and_left = sheet_active.find('.excel-main-table-body-and-left');
    
            var delta = (event.originalEvent.wheelDelta && (event.originalEvent.wheelDelta > 0 ? 1 : -1)) || (event.originalEvent.detail && (event.originalEvent.detail > 0 ? -1 : 1));
            if (delta > 0) {
                delta = -1;
            } else if (delta < 0) {
                delta = 1;
            } else {
                delta = 0;
            }
    
            top = 30 * delta;
            temp = body_and_left[0].scrollHeight - body_and_left.height();
            move = temp - body_and_left.scrollTop();
            if(top > 0)
            {
                if(top > move)
                {
                    insertTable(0, present(' tr').length-1, 'iDown');
                    top = move;
                }
            }
            if(top < 0)
            {
                if(top + body_and_left.scrollTop() < 0)
                {
                    top = -body_and_left.scrollTop();
                }
            }
            body_and_left.scrollTop(body_and_left.scrollTop() + top);   
            cha = that.find('.excel-main-table-scroll').height() - that.find('.excel-main-table-scrollbar').height();
            that.find('.excel-main-table-scrollbar').css('top', body_and_left.scrollTop() / temp * cha);
        });
    }

    // 初始化事件绑定
    function initBind()
    {
        that.resize(initTable);

        // 表格标题操作事件
        excelBottomSheet();
        // 底部横向滚动条事件
        excelScroll();
        // 表格内容操作事件
        excelMainTableBody();
    }

    // table_body.append('<div class="excel-main-table-cheek" style="width: '+td_width+'px; height: '+td_height+'px;"></div>');
    // that.find('input[name="excel-head-input"]').val(temp_table_body.find("tr:first td:first").html());

    /* 下面是功能已完成的代码 */

    // 格式化表格
    function initTable()
    {        
        // 获取当前页
        let sheet_active = that.find('.excel-main-table-data.active');
        if(sheet_active.length == 0)
        {
            that.find('.excel-bottom-sheet a.excel-bottom-sheet-oper[mold="add"]').click();
            return ;
        }
        // 获取表格横坐标
        let table_head = sheet_active.find('.excel-main-table-head');
        // 获取表格纵坐标
        let table_left = sheet_active.find('.excel-main-table-left');
        // 获取表格主体
        let table_body = sheet_active.find('.excel-main-table-body');
        // 获取表格主体和纵坐标
        let body_and_left = sheet_active.find('.excel-main-table-body-and-left');

        // 获取横坐标表格
        let temp_table_head = $('<table><tr></tr></table>');
        // 获取纵坐标表格
        let temp_table_left = table_left.find("table");
        temp_table_left = temp_table_left.length == 0 ? $('<table></table>') : temp_table_left.first().clone(false);
        // 获取主体表格
        let temp_table_body = table_body.find("table");
        temp_table_body = temp_table_body.length == 0 ? $('<table></table>') : temp_table_body.first().clone(false);

        // 这里声明不会更改值的变量
        let temp_old_cols, temp_old_rows, temp_old_width, temp_old_height,
            temp_new_cols, temp_new_rows, temp_count_cols, temp_count_rows;

        // 这里声明公共变量
        let temp_td, temp_tr, temp_width, temp_table;

        // 获取原表格第一行所有td
        temp_td = temp_table_body.find('tr:first td');
        // 获取原表格列数
        temp_old_cols = temp_td.length;
        // 获取原表格所有行
        temp_tr = temp_table_body.find('tr');
        // 获取原表格行数
        temp_old_rows = temp_tr.length;

        // 设置原表格宽度
        temp_old_width = 0;
        for(let i=0; i< temp_old_cols; i++)
        {
            temp_width = temp_td.eq(i).width();
            if(temp_width == 0)
            {
                temp_td.eq(i).css('width',td_width);
                temp_width = td_width;
            }
            temp_old_width += temp_width;
        }
        // 设置原表格高度
        temp_old_height = temp_old_rows * td_height;
        
        // 设置需要新增的列数
        temp_new_cols = Math.ceil((table_body.width() - temp_old_width) / td_width);
        // 设置需要新增的行数
        temp_new_rows = Math.ceil((body_and_left.height() - temp_old_height) / td_height);

        // 设置总列长
        temp_count_cols = temp_new_cols + temp_old_cols;
        // 设置总行长
        temp_count_rows = temp_new_rows + temp_old_rows;

        // 补齐原表列
        if(temp_new_cols > 0)
        {
            for(let i=0; i<temp_old_rows; i++)
            {
                temp_tr.eq(i).height(td_height);
                temp_tr.eq(i).append('<td></td>'.repeat(temp_new_cols));
            }
        }

        // 添加新行
        if(temp_new_rows > 0)
        {
            temp_table_body.append(('<tr style="height:'+td_height+'px;">' + '<td></td>'.repeat(temp_count_cols) +'</tr>').repeat(temp_new_rows));
        }

        // 设置纵坐标
        temp_tr = temp_table_left.find('tr');
        for(let i=temp_tr.length; i<temp_count_rows; i++)
        {
            temp_table_left.append('<tr style="height: '+td_height+'px;"><td>' + (i+1) + '</td></tr>');
        }

        // 获取原表格第一行所有td
        temp_td = temp_table_body.find('tr:first td');
        for(let i=0; i<temp_td.length; i++)
        {
            temp_width = temp_td.eq(i).width();
            if(temp_width == 0)
            {
                temp_td.eq(i).css('width',td_width);
                temp_width = td_width;
            }
            temp_table_head.find('tr').append('<td style="width: '+temp_width+'px;">' + getColumn(i+1) + '</td>')
        }

        temp_width = temp_old_width + td_width * temp_new_cols;
        temp_table_head.css('width', temp_width);
        temp_table_body.css('width', temp_width);

        table_head.html(temp_table_head);
        table_left.html(temp_table_left);

        temp_table = table_body.find("table:first");
        if(temp_table.length > 0)
        {
            temp_table.replaceWith(temp_table_body);
        }else{
            table_body.html(temp_table_body);
        }
        
        if(table_body.find('.excel-main-table-cheek').length == 0)
        {
            temp_table_body.find("tr:first td:first").click();
        }

        initScroll();
    }

    // 初始化滚动条
    function initScroll()
    {
        let sheet_active = that.find('.excel-main-table-data.active');
        let table_left = sheet_active.find('.excel-main-table-left');
        let table_body = sheet_active.find('.excel-main-table-body');
        let temp_table_body = table_body.find("table").first();
        let temp_table_left = table_left.find("table").first();
        let body_and_left = sheet_active.find('.excel-main-table-body-and-left');

        let width_show = temp_table_body.width();
        let width_body = table_body.width();
        let height_show = temp_table_left.height();
        let height_body = body_and_left.height();
        
        that.find('.excel-bottom-scrollbar').css('width', (width_body / width_show * 100) +'%');
        that.find('.excel-main-table-scrollbar').css('height', (height_body / height_show * 100) +'%');
    }

    // 底部横向滚动条事件
    function excelScroll()
    {
        // 横向滚动条
        that.find('.excel-bottom-scrollbar').mousedown(function(event){
            let start = event.pageX;
            let move = that.find('.excel-bottom-scroll').width() - $(this).width();
            let sheet_active = that.find('.excel-main-table-data.active');
            let table_head = sheet_active.find('.excel-main-table-head');
            let table_body = sheet_active.find('.excel-main-table-body');
            $(document).mousemove(function(event) {
                let last = event.pageX - start;
                let left = parseInt(that.find('.excel-bottom-scrollbar').css('left')) + last;
                left = left < 0 ? 0 : move < left ? move : left;
                start = event.pageX;
                table_head.scrollLeft((table_head[0].scrollWidth - table_head.width()) * (left / move));
                table_body.scrollLeft((table_body[0].scrollWidth - table_body.width()) * (left / move));
                that.find('.excel-bottom-scrollbar').css('left', left);
            });
            $(document).mouseup(function(){
                $(document).off('mousemove');
            });
        });
        // 纵向滚动条
        that.find('.excel-main-table-scrollbar').mousedown(function(event){
            let start = event.pageY;
            let move = that.find('.excel-main-table-scroll').height() - $(this).height();
            let sheet_active = that.find('.excel-main-table-data.active');
            let body_and_left = sheet_active.find('.excel-main-table-body-and-left');
    
            $(document).mousemove(function(event) {
                let last = event.pageY - start;
                let top = parseInt(that.find('.excel-main-table-scrollbar').css('top')) + last;
                top = top < 0 ? 0 : move < top ? move : top;
                start = event.pageY;
                body_and_left.scrollTop((body_and_left[0].scrollHeight - body_and_left.height()) * (top / move));
                that.find('.excel-main-table-scrollbar').css('top', top);
            });
            $(document).mouseup(function(){
                $(document).off('mousemove');
            });
        });
    }

    // 表格标题操作事件
    function excelBottomSheet()
    {
        let present = function(child = ''){
            return that.find('.excel-bottom-sheet' + child);
        };
        let sheetOper = function(mold=''){
            return that.find('.excel-bottom-sheet a.excel-bottom-sheet-oper'+mold);
        };
        let table = function(child = ''){
            return that.find('.excel-main-table' + child);
        }

        // 尺寸大小改变事件
        present('-tab-title').resize(resize);
        function resize(){
            let temp_bool = present().width() - 62 - present('-tab-title').width() < 0;
            sheetOper('[mold="right"]').attr('disabled', !temp_bool);
            parseInt(present('-tab-title').css('marginLeft')) < 0;
            sheetOper('[mold="left"]').attr('disabled', !temp_bool);
        }
        resize();

        // 按钮点击事件
        present().off('click', 'a.excel-bottom-sheet-oper').on('click', 'a.excel-bottom-sheet-oper', function(){
            let mold = $(this).attr('mold'), tab = present('-tab'), ul = present('-tab-title'), temp, margin_left;
            if($(this).attr('disabled') != 'disabled')
            {
                switch(mold)
                {
                    case 'left':
                        margin_left = parseInt(ul.css('marginLeft')) + 20;
                        if(margin_left >= 0)
                        {
                            $(this).attr('disabled', true);
                            margin_left = 0;
                        }
                        sheetOper('[mold="right"]').attr('disabled', false);
                        ul.css('marginLeft', margin_left);
                        break;
                    case 'right':
                        temp = tab.width() - ul.width();
                        margin_left = parseInt(ul.css('marginLeft')) - 20;
                        if(temp >= margin_left)
                        {
                            $(this).attr('disabled', true);
                            margin_left = temp;
                        }
                        sheetOper('[mold="left"]').attr('disabled', false);
                        ul.css('marginLeft', margin_left);
                        break;
                    case 'add':
                        $(this).attr('disabled', true);
                        let id = 'new' + (Date.parse(new Date()) + Math.round(Math.random()*10));
                        present('-tab-title li').removeClass('active');
                        present('-tab-title').append('<li data-i="'+id+'">new Sheet</li>');
                        present('-tab-title li[data-i="'+id+'"]').addClass('active');
                        let html = '<div data-i="'+id+'" class="excel-main-table-data"><div class="excel-main-table-select"></div><div class="excel-main-table-head"></div><div class="excel-main-table-body-and-left"><div class="excel-main-table-left"></div><div class="excel-main-table-body"></div><div class="clear"></div></div></div>';
                        table().append(html);
                        table('-data.active').removeClass('active');
                        table('-data[data-i="'+id+'"]').addClass('active');
                        temp1 = tab.width() - ul.width();
                        if(temp1 < 0)
                        {
                            ul.css('marginLeft', temp1);
                        }
                        that.resize();
                        setTimeout(function(){
                            sheetOper('[mold="add"]').attr('disabled', false);
                        }, 500);
                        break;
                }
            }
        });

        // 表格标题单击事件
        present().off('click', '.excel-bottom-sheet-tab-title li').on('click', '.excel-bottom-sheet-tab-title li', function(){
            let id = $(this).attr('data-i');
            if(!$(this).hasClass('active'))
            {
                present('-tab-title li').removeClass('active');
                present('-tab-title li[data-i="'+id+'"]').addClass('active');
                table('-data.active').removeClass('active');
                table('-data[data-i="'+id+'"]').addClass('active');
                that.resize();
            }
        });

        // 表格标题双击事件
        present().off('dblclick', '.excel-bottom-sheet-tab-title li').on('dblclick', '.excel-bottom-sheet-tab-title li', function(){
            $(this).attr('contentEditable', true);
            $(this).focus();
            if (document.body.createTextRange) {
                var range = document.body.createTextRange();
                range.moveToElementText(this);
                range.select();
            } else if (window.getSelection) {
                var selection = window.getSelection();
                var range = document.createRange();
                range.selectNodeContents(this);
                selection.removeAllRanges();
                selection.addRange(range);
            }
            $(this).blur(function(){
                $(this).attr('contentEditable', false);
            });
        });

        // 右键菜单事件
        present().off('contextmenu', 'li').on('contextmenu', 'li', function(e) {
            let temp_this = $(this);
            that.append('<div class="contextmenu-sheet"><div data="del"><span class="icon">&#xe60e;</span>删除</div><div data="edit"><span class="icon">&#xe620;</span>重命名</div></div>');
			// 获取窗口尺寸
			var winWidth = $(document).width();
			var winHeight = $(document).height();
			// 鼠标点击位置坐标
			var mouseX = e.pageX;
			var mouseY = e.pageY;
			// ul标签的宽高
			var menuWidth = $(".contextmenu-sheet").width();
			var menuHeight = $(".contextmenu-sheet").height();
			// 最小边缘margin(具体窗口边缘最小的距离)
			var minEdgeMargin = 10;
			// 以下判断用于检测ul标签出现的地方是否超出窗口范围
			// 第一种情况：右下角超出窗口
			if(mouseX + menuWidth + minEdgeMargin >= winWidth &&
				mouseY + menuHeight + minEdgeMargin >= winHeight) {
				menuLeft = mouseX - menuWidth - minEdgeMargin + "px";
				menuTop = mouseY - menuHeight - minEdgeMargin + "px";
			}
			// 第二种情况：右边超出窗口
			else if(mouseX + menuWidth + minEdgeMargin >= winWidth) {
				menuLeft = mouseX - menuWidth - minEdgeMargin + "px";
				menuTop = mouseY + minEdgeMargin + "px";
			}
			// 第三种情况：下边超出窗口
			else if(mouseY + menuHeight + minEdgeMargin >= winHeight) {
				menuLeft = mouseX + minEdgeMargin + "px";
				menuTop = mouseY - menuHeight - minEdgeMargin + "px";
			}
			// 其他情况：未超出窗口
			else {
				menuLeft = mouseX + minEdgeMargin + "px";
				menuTop = mouseY + minEdgeMargin + "px";
			};
            // ul菜单出现
            that.find('.contextmenu-sheet').css({
				"left": menuLeft,
				"top": menuTop
            }).show();
            // 点击之后，右键菜单隐藏
            that.off('click', '.contextmenu-sheet div').on('click', '.contextmenu-sheet div', function(){
                let mold = $(this).attr('data');
                if(mold == 'edit')
                {
                    temp_this.dblclick();
                }else if(mold == 'del'){
                    id = temp_this.attr('data-i');
                    temp_this.remove();
                    table('-data[data-i="'+id+'"]').remove();
                    if(present(' li').length < 1)
                    {
                        sheetOper('[mold="add"]').click();
                    }else{
                        present(' li').eq(0).click();
                    }
                }
                // console.log(temp_this);
            });
            $(document).click(function() {
                $(".contextmenu-sheet").remove();
            });
			// 阻止浏览器默认的右键菜单事件
			return false;
        });
    }

    // 获取列名
    function getColumn(num){
        var str = '';
        if(num > 0) {
            if(num >= 1 && num <= 26) {
                str = String.fromCharCode(64 + parseInt(num));
            } else {
                while(num > 26) {
                    var count = parseInt(num/26);
                    var remainder = num % 26;
                    if(remainder == 0) {
                        remainder = 26;
                        count--;
                        str = String.fromCharCode(64 + parseInt(remainder)) + str;
                    } else {
                        str = String.fromCharCode(64 + parseInt(remainder)) + str;
                    }
                    num = count;
                }
                str = String.fromCharCode(64 + parseInt(num)) + str;
            }
        }
        return str;
    }
})(jQuery);

// 元素resize事件
(function($, h, c) {
    var a = $([]), e = $.resize = $.extend($.resize, {}), i, k = "setTimeout", j = "resize", d = j + "-special-event", b = "delay", f = "throttleWindow";
    e[b] = 350;
    e[f] = true;
    $.event.special[j] = {
        setup : function() {
            if (!e[f] && this[k]) {
                return false
            }
            var l = $(this);
            a = a.add(l);
            $.data(this, d, {
                w : l.width(),
                h : l.height()
            });
            if (a.length === 1) {
                g()
            }
        },
        teardown : function() {
            if (!e[f] && this[k]) {
                return false
            }
            var l = $(this);
            a = a.not(l);
            l.removeData(d);
            if (!a.length) {
                clearTimeout(i)
            }
        },
        add : function(l) {
            if (!e[f] && this[k]) {
                return false
            }
            var n;
            function m(s, o, p) {
                var q = $(this), r = $.data(this, d);
                r.w = o !== c ? o : q.width();
                r.h = p !== c ? p : q.height();
                n.apply(this, arguments)
            }
            if ($.isFunction(l)) {
                n = l;
                return m
            } else {
                n = l.handler;
                l.handler = m
            }
        }
    };
    function g() {
        i = h[k](function() {
            a.each(function() {
                var n = $(this), m = n.width(), l = n.height(), o = $
                        .data(this, d);
                if (m !== o.w || l !== o.h) {
                    n.trigger(j, [ o.w = m, o.h = l ])
                }
            });
            g()
        }, e[b])
    }
})(jQuery, this);
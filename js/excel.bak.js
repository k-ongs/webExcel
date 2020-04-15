(function ($) {
    let cellNum, rowNum, cellWidth, cellHeight, excel, option;
    // 注册JQ函数
    $.fn.extend({
        excel: function (options) {
            excel = this;
            option = $.extend({}, options);
            // 设置列宽
            cellWidth = option.cellWidth ? option.cellWidth : 80;
            // 设置行高
            cellHeight = option.cellHeight ? option.cellHeight : 30;
            // 初始化excel
            initExcel();
            // 绑定事件
            initBind();
            // 初始化表格
            initTable();
        }
    });

    // 初始化事件绑定
    function initExcel()
    {
        excel.addClass('excel');
        if(excel.find('.excel-sheet-tab').width() - excel.find('.excel-sheet-tab-title').width() < 0)
        {
            excel.find('a.excel-sheet-oper[mold="right"]').attr('disabled', false);
        }
    }

    // 初始化事件绑定
    function initBind()
    {
        excel.on('click', 'a.excel-sheet-oper', excelSheetOper);
        excel.find('.excel-table').on('scroll', scrollHandle);
        excel.resize(formatTable);
        excel.find('.excel-bottom-scrollbar').mousedown(function(event){
            let start = event.pageX;
            let move = excel.find('.excel-bottom-scroll').width() - $(this).width();
            // (excel.find('.excel-bottom-scroll').width() / 600)
            $(document).mousemove(function(event) {
                let last = event.pageX - start;
                let left = parseInt(excel.find('.excel-bottom-scrollbar').css('left')) + last;
                left = left < 0 ? 0 : move < left ? move : left;
                start = event.pageX;
                excel.find('.excel-table').scrollLeft((excel.find('.excel-table')[0].scrollWidth - excel.find('.excel-table').width()) * (left / move));
                excel.find('.excel-bottom-scrollbar').css('left', left);
            });
            $(document).mouseup(function(){
                $(document).off('mousemove');
            });
        });
    }

    function scrollHandle(e) {
        let scrollLeft = this.scrollLeft;
        let scrollTop = this.scrollTop;
        let d = $(this).data('slt');
        let sheet_active = excel.find('.excel-sheet-data.active');

        if (scrollLeft != (d == undefined ? 0 : d.sl)) {
            sheet_active.find("table tr").each(function(){
                $(this).children("td:first").css('transform', 'translateX(' + scrollLeft + 'px)');
            })
        }
        if (scrollTop != (d == undefined ? 0 : d.st)) {
            sheet_active.find("table tr:first td").css('transform', 'translateY(' + scrollTop + 'px)');
        }
        let first = sheet_active.find("table tr:first td:first");
        let first_list = first.css("transform").replace(/[^0-9\-,]/g,'').split(',');
        scrollLeft = (scrollLeft != (d == undefined ? 0 : d.sl)) ? first_list[4] : scrollLeft;
        scrollTop = (scrollTop != (d == undefined ? 0 : d.st)) ? first_list[5] : scrollTop;
        first.css('transform', 'translate(' + scrollLeft + 'px,' + scrollTop + 'px)');
        $(this).data('slt', {sl: scrollLeft, st: scrollTop});
    }

    function excelSheetOper()
    {
        let mold, tab, ul, temp1, temp2;
        if($(this).attr('disabled') != 'disabled')
        {
            tab = excel.find('.excel-sheet-tab');
            ul = excel.find('.excel-sheet-tab-title');
            mold = $(this).attr('mold');
            switch(mold)
            {
                case 'left':
                    temp1 = parseInt(ul.css('marginLeft')) + 20;
                    if(temp1 >= 0)
                    {
                        $(this).attr('disabled', true);
                        temp1 = 0;
                    }
                    excel.find('a.excel-sheet-oper[mold="right"]').attr('disabled', false);
                    ul.css('marginLeft', temp1);
                    break;
                case 'right':
                    temp1 = tab.width() - ul.width();
                    temp2 = parseInt(ul.css('marginLeft')) - 20;
                    if(temp1 >= temp2)
                    {
                        $(this).attr('disabled', true);
                        temp2 = temp1;
                    }
                    excel.find('a.excel-sheet-oper[mold="left"]').attr('disabled', false);
                    ul.css('marginLeft', temp2);
                    break;
                case 'add':
                    // 
                    break;
            }
        }
    }

    // 格式化表格
    function formatTable()
    {
        let tem_cellNum =  Math.ceil((excel.width() - 40) / cellWidth);
        let tem_rowNum =  Math.ceil((excel.height() - 80) / cellHeight);
        console.log(tem_rowNum);
        if(tem_cellNum > cellNum)
        {
            temp = '';
            for(let i = cellNum; i < tem_cellNum ; i++)
            {
                temp += '<td style="width: '+cellWidth+'px;">'+(getColumn(i+1))+'</td>';
            }
            excel.find("table tr:first").append(temp);
            excel.find("table tr:not(:first)").append('<td></td>'.repeat(tem_cellNum - cellNum));
            cellNum = tem_cellNum;
        }
        if(tem_rowNum > rowNum)
        {
            excel.find("table").append(('<tr style="height: '+cellHeight+'px;">' + '<td></td>'.repeat(cellNum) +'</tr>').repeat(tem_rowNum - rowNum));
            for(let i = rowNum +1; i < tem_rowNum+1; i++)
            {
                excel.find("table tr").eq(i).prepend('<td>' + i + '</td>');
            }
            rowNum = tem_rowNum;
        }else{
            for(let i = rowNum; i > tem_rowNum; i--)
            {
                if('<td>'+i+'</td>'+'<td></td>'.repeat(cellNum) == excel.find("table tr").eq(i).html())
                {
                    rowNum--;
                    excel.find("table tr").eq(i).remove();
                }else{
                    break;
                }
            }
        }
        excel.find('.excel-bottom-scrollbar').css('width', (excel.find('.excel-table').width() / excel.find('.excel-table')[0].scrollWidth * 100)+'%');
    }

    // 初始化表格
    function initTable()
    {
        let sheet_active,table, temp, tr, td;
        sheet_active = excel.find('.excel-sheet-data.active');

        cellNum = Math.ceil((excel.width() - 40) / cellWidth);
        rowNum = Math.ceil((excel.height() - 80) / cellHeight);

        table = sheet_active.find("table").first().clone(false);

        tr = table.find("tr");
        for(let j = tr.children('td').length; j <= cellNum ; j++)
        {
            tr.append('<td></td>');
        }
        temp = '';
        for(let i = 0; i < cellNum ; i++)
        {
            temp += '<td>'+(getColumn(i+1))+'</td>';
        }
        table.prepend('<tr>'+temp+'</tr>');
        table.append(('<tr>' + '<td></td>'.repeat(cellNum) +'</tr>').repeat(rowNum - tr.length));
        tr = table.find("tr");
        tr.eq(0).prepend('<td></td>');
        for(let i = 1; i < tr.length; i++)
        {
            tr.eq(i).prepend('<td>' + i + '</td>');
        }
        table.find("tr").height(cellHeight);
        table.find("tr").first().children('td').width(cellWidth);
        sheet_active.find("table").first().replaceWith(table);
        
        table = sheet_active.find("table").first();
        excel.find('.excel-bottom-scrollbar').css('width', (excel.find('.excel-table').width() / excel.find('.excel-table')[0].scrollWidth * 100)+'%');
        // $('.excel-table').scrollLeft(100);
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

(function($, h, c) {
    var a = $([]), e = $.resize = $.extend($.resize, {}), i, k = "setTimeout", j = "resize", d = j
            + "-special-event", b = "delay", f = "throttleWindow";
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
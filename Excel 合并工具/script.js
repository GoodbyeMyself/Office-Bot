// Excel文件汇总工具 - JavaScript模块
const { createApp } = Vue;

// 创建Vue应用实例
const ExcelMergerApp = createApp({
    data() {
        return {
            mainFile: null,
            mergeFiles: [],
            loading: false,
            loadingText: '正在处理文件...',
            errorMessage: '',
            successMessage: '',
            fileIdCounter: 0
        }
    },
    computed: {
        canMerge() {
            return this.mainFile && this.mergeFiles.length > 0;
        }
    },
    methods: {
        // 处理主表文件选择
        async handleMainFileSelect(file) {
            if (!file) return;

            this.clearMessages();
            this.loading = true;
            this.loadingText = '正在读取主表文件...';

            try {
                const data = await this.readExcelFile(file);
                this.mainFile = {
                    name: file.name,
                    size: file.size,
                    data: data,
                    file: file
                };
                ElMessage.success('主表文件读取成功！');
            } catch (error) {
                ElMessage.error('读取主表文件失败：' + error.message);
            } finally {
                this.loading = false;
            }
        },

        // 处理待合并文件选择
        async handleMergeFileSelect(fileList) {
            if (!fileList || fileList.length === 0) return;

            this.clearMessages();
            this.loading = true;
            this.loadingText = `正在读取 ${fileList.length} 个文件...`;

            try {
                let addedCount = 0;
                for (const file of fileList) {
                    // 检查文件是否已经添加过
                    if (this.mergeFiles.some(f => f.name === file.name && f.size === file.size)) {
                        console.warn(`文件 ${file.name} 已存在，跳过...`);
                        continue;
                    }

                    const data = await this.readExcelFile(file);
                    this.mergeFiles.push({
                        id: ++this.fileIdCounter,
                        name: file.name,
                        size: file.size,
                        data: data,
                        file: file
                    });
                    addedCount++;
                }
                
                if (addedCount > 0) {
                    ElMessage.success(`成功添加 ${addedCount} 个文件！`);
                }
            } catch (error) {
                ElMessage.error('读取文件失败：' + error.message);
            } finally {
                this.loading = false;
            }
        },

        // 移除待合并文件
        removeMergeFile(index) {
            this.mergeFiles.splice(index, 1);
            ElMessage.info('文件已移除');
        },

        // 读取Excel文件
        readExcelFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: 'array'});
                        const sheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                        resolve(jsonData);
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = () => reject(new Error('文件读取失败'));
                reader.readAsArrayBuffer(file);
            });
        },

        // 执行合并 - 直接在主表上合并
        async performMerge() {
            if (!this.canMerge) return;

            this.clearMessages();
            this.loading = true;
            this.loadingText = '正在直接合并到主表文件...';

            try {
                // 直接修改主表数据
                const originalRowCount = this.mainFile.data.length;
                
                // 合并所有待合并文件的数据（跳过表头）
                this.mergeFiles.forEach(file => {
                    const dataRows = file.data.slice(1);
                    this.mainFile.data.push(...dataRows);
                });

                const newRowCount = this.mainFile.data.length;
                const addedRows = newRowCount - originalRowCount;

                // 立即保存修改后的主表文件
                await this.saveMainFileDirectly();

                ElMessage.success(`合并完成！已直接添加 ${addedRows} 行数据到主表文件中`);
                
                // 清空待合并文件列表，因为已经合并完成
                this.mergeFiles = [];
                
            } catch (error) {
                ElMessage.error('合并失败：' + error.message);
            } finally {
                this.loading = false;
            }
        },

        // 直接保存主表文件
        async saveMainFileDirectly() {
            try {
                const ws = XLSX.utils.aoa_to_sheet(this.mainFile.data);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "合并数据");
                
                // 使用原文件名保存
                const fileName = this.mainFile.name;
                XLSX.writeFile(wb, fileName);
                
                ElMessage.success(`主表文件已更新：${fileName}`);
            } catch (error) {
                throw new Error('保存主表文件失败：' + error.message);
            }
        },



        // 清除消息
        clearMessages() {
            this.errorMessage = '';
            this.successMessage = '';
        },

        // Element Plus 上传组件的钩子函数
        beforeMainUpload(file) {
            const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                           file.type === 'application/vnd.ms-excel' ||
                           file.name.endsWith('.xlsx') || 
                           file.name.endsWith('.xls');
            
            if (!isExcel) {
                ElMessage.error('只能上传Excel文件(.xlsx/.xls)！');
                return false;
            }
            
            // 直接处理文件，不需要实际上传
            this.handleMainFileSelect(file);
            return false; // 阻止上传
        },

        // 主文件上传成功回调（实际不会用到，因为阻止了上传）
        handleMainUploadSuccess(response, file) {
            // 不需要实现
        },

        // 待合并文件上传前验证
        beforeMergeUpload(file) {
            const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                           file.type === 'application/vnd.ms-excel' ||
                           file.name.endsWith('.xlsx') || 
                           file.name.endsWith('.xls');
            
            if (!isExcel) {
                ElMessage.error('只能上传Excel文件(.xlsx/.xls)！');
                return false;
            }

            // 检查文件是否已经添加过
            if (this.mergeFiles.some(f => f.name === file.name && f.size === file.size)) {
                ElMessage.warning(`文件 ${file.name} 已存在，跳过...`);
                return false;
            }
            
            // 直接处理文件
            this.handleSingleMergeFile(file);
            return false; // 阻止上传
        },

        // 处理单个待合并文件
        async handleSingleMergeFile(file) {
            this.loading = true;
            this.loadingText = `正在读取文件 ${file.name}...`;

            try {
                const data = await this.readExcelFile(file);
                this.mergeFiles.push({
                    id: ++this.fileIdCounter,
                    name: file.name,
                    size: file.size,
                    data: data,
                    file: file
                });
                ElMessage.success(`成功添加文件：${file.name}`);
            } catch (error) {
                ElMessage.error(`读取文件 ${file.name} 失败：${error.message}`);
            } finally {
                this.loading = false;
            }
        },

        // 待合并文件上传成功回调（实际不会用到）
        handleMergeUploadSuccess(response, file, fileList) {
            // 不需要实现
        },

        // 处理上传错误
        handleUploadError(error, file) {
            ElMessage.error(`文件 ${file.name} 上传失败`);
        },

        // 文件超出限制时的钩子（已取消限制）
        handleExceed(files, fileList) {
            // 不再限制文件数量
        },

        // 移除文件
        handleRemove(file, fileList) {
            // 找到对应的文件并移除
            const fileName = file.name;
            const index = this.mergeFiles.findIndex(f => f.name === fileName);
            if (index > -1) {
                this.removeMergeFile(index);
            }
        }
    }
});

// 导出应用实例（供HTML使用）
window.ExcelMergerApp = ExcelMergerApp;
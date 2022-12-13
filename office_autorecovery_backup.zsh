#!/bin/zsh
# 备份Office三件套的自动恢复目录    Back up the auto recovery directory of Office
# 配合launchd使用   Please use it with launchd

set -eu

# 设置变量  Set up variables
# Office自动恢复目录    Office AutoRecovery directories
word_autorecovery_dir="/Users/$(whoami)/Library/Containers/com.microsoft.Word/Data/Library/Preferences/AutoRecovery"
excel_autorecovery_dir="/Users/$(whoami)/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft"
powerpoint_autorecovery_dir="/Users/$(whoami)/Library/Containers/com.microsoft.Powerpoint/Data/Library/Preferences/AutoRecovery"
# 备份目录，有需要的话自行更改  Backup directories, change as you need
word_backup_dir="/Users/$(whoami)/Library/CloudStorage/OneDrive-个人/备份/Office自动恢复/Word"
excel_backup_dir="/Users/$(whoami)/Library/CloudStorage/OneDrive-个人/备份/Office自动恢复/Excel"
powerpoint_backup_dir="/Users/$(whoami)/Library/CloudStorage/OneDrive-个人/备份/Office自动恢复/PowerPoint"

# 计算现有最新备份文件的sha256  Compute sha256 of the latest backup file
word_latest_filename=$(/bin/ls -t ${word_backup_dir} | head -n1 | awk '{print $0}')
excel_latest_filename=$(/bin/ls -t ${excel_backup_dir} | head -n1 | awk '{print $0}')
powerpoint_latest_filename=$(/bin/ls -t ${powerpoint_backup_dir} | head -n1 | awk '{print $0}')

word_latest_sha256=""
excel_latest_sha256=""
powerpoint_latest_sha256=""
if [[ ! ${word_latest_filename} == "" ]]
then
    word_latest_sha256=$(sha256sum ${word_backup_dir}/${word_latest_filename})
fi
if [[ ! ${excel_latest_filename} == "" ]]
then
    excel_latest_sha256=$(sha256sum ${excel_backup_dir}/${excel_latest_filename})
fi
if [[ ! ${powerpoint_latest_filename} == "" ]]
then
    powerpoint_latest_sha256=$(sha256sum ${powerpoint_backup_dir}/${powerpoint_latest_filename})
fi

# 压缩文件  Compress auto recovery files
tar czf ${word_backup_dir}/$(date +%Y%m%d%H%M%S).tar.gz --directory ${word_autorecovery_dir} .
tar czf ${excel_backup_dir}/$(date +%Y%m%d%H%M%S).tar.gz --directory ${excel_autorecovery_dir} .
tar czf ${powerpoint_backup_dir}/$(date +%Y%m%d%H%M%S).tar.gz --directory ${powerpoint_autorecovery_dir} .

# 验证是否重复  Verify if the backup file is duplicated
word_new_latest_filename=$(/bin/ls -t ${word_backup_dir} | head -n1 | awk '{print $0}')
excel_new_latest_filename=$(/bin/ls -t ${excel_backup_dir} | head -n1 | awk '{print $0}')
powerpoint_new_latest_filename=$(/bin/ls -t ${powerpoint_backup_dir} | head -n1 | awk '{print $0}')

word_new_latest_sha256=$(sha256sum ${word_backup_dir}/${word_new_latest_filename})
excel_new_latest_sha256=$(sha256sum ${excel_backup_dir}/${excel_new_latest_filename})
powerpoint_new_latest_sha256=$(sha256sum ${powerpoint_backup_dir}/${powerpoint_new_latest_filename})

if [[ ${word_new_latest_sha256} == ${word_latest_sha256} ]]
then
    rm ${word_backup_dir}/${word_new_latest_filename}
fi
if [[ ${excel_new_latest_sha256} == ${excel_latest_sha256} ]]
then
    rm ${excel_backup_dir}/${excel_new_latest_filename}
fi
if [[ ${powerpoint_new_latest_sha256} == ${powerpoint_latest_sha256} ]]
then
    rm ${powerpoint_backup_dir}/${powerpoint_new_latest_filename}
fi
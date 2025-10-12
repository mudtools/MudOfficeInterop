using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentProtectionAndSecuritySample
{
    /// <summary>
    /// 权限管理器类
    /// </summary>
    public class PermissionManager
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public PermissionManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 检查权限管理是否可用
        /// </summary>
        /// <returns>是否可用</returns>
        public bool IsPermissionManagementAvailable()
        {
            try
            {
                bool isEnabled = _document.Permission.Enabled;
                Console.WriteLine($"权限管理是否可用: {isEnabled}");
                return isEnabled;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"检查权限管理可用性时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 启用权限管理
        /// </summary>
        /// <returns>是否启用成功</returns>
        public bool EnablePermissionManagement()
        {
            try
            {
                _document.Permission.Enabled = true;
                Console.WriteLine("权限管理已启用");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"启用权限管理时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 禁用权限管理
        /// </summary>
        /// <returns>是否禁用成功</returns>
        public bool DisablePermissionManagement()
        {
            try
            {
                _document.Permission.Enabled = false;
                Console.WriteLine("权限管理已禁用");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"禁用权限管理时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加用户权限
        /// </summary>
        /// <param name="userEmail">用户邮箱</param>
        /// <param name="permissions">权限列表</param>
        /// <param name="expirationDate">过期日期</param>
        /// <returns>用户权限对象</returns>
        public IWordUserPermission AddUserPermission(
            string userEmail, 
            List<MsoPermission> permissions, 
            DateTime? expirationDate = null)
        {
            try
            {
                // 计算权限值
                int permissionValue = 0;
                foreach (var permission in permissions)
                {
                    permissionValue += (int)permission;
                }

                // 添加用户权限
                var userPermission = _document.Permission.Add(userEmail, permissionValue);

                // 设置过期日期
                if (expirationDate.HasValue)
                {
                    userPermission.ExpirationDate = expirationDate.Value;
                }

                Console.WriteLine($"用户权限已添加: {userEmail}");
                return userPermission;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加用户权限时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 移除用户权限
        /// </summary>
        /// <param name="userEmail">用户邮箱</param>
        /// <returns>是否移除成功</returns>
        public bool RemoveUserPermission(string userEmail)
        {
            try
            {
                _document.Permission.Remove(userEmail);
                Console.WriteLine($"用户权限已移除: {userEmail}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移除用户权限时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 应用权限策略
        /// </summary>
        /// <param name="policyPath">策略文件路径</param>
        /// <returns>是否应用成功</returns>
        public bool ApplyPermissionPolicy(string policyPath)
        {
            try
            {
                _document.Permission.ApplyPolicy(policyPath);
                Console.WriteLine($"权限策略已应用: {policyPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用权限策略时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取所有用户权限信息
        /// </summary>
        /// <returns>用户权限信息列表</returns>
        public List<UserPermissionInfo> GetAllUserPermissions()
        {
            var permissions = new List<UserPermissionInfo>();

            try
            {
                for (int i = 1; i <= _document.Permission.Count; i++)
                {
                    var userPermission = _document.Permission.Item(i);
                    var permissionInfo = new UserPermissionInfo
                    {
                        UserId = userPermission.UserId,
                        PermissionValue = userPermission.Permission,
                        ExpirationDate = userPermission.ExpirationDate
                    };
                    permissions.Add(permissionInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取用户权限信息时出错: {ex.Message}");
            }

            return permissions;
        }

        /// <summary>
        /// 创建标准企业权限策略
        /// </summary>
        /// <param name="allowedUsers">允许的用户列表</param>
        /// <param name="defaultExpirationDays">默认过期天数</param>
        /// <returns>是否创建成功</returns>
        public bool CreateStandardCorporatePolicy(List<string> allowedUsers, int defaultExpirationDays = 30)
        {
            try
            {
                // 启用权限管理
                if (!EnablePermissionManagement())
                {
                    return false;
                }

                // 为每个用户添加权限
                foreach (var user in allowedUsers)
                {
                    var permissions = new List<MsoPermission>
                    {
                        MsoPermission.msoPermissionRead,
                        MsoPermission.msoPermissionEdit
                    };

                    var expirationDate = DateTime.Now.AddDays(defaultExpirationDays);

                    AddUserPermission(user, permissions, expirationDate);
                }

                Console.WriteLine($"标准企业权限策略已创建，包含 {allowedUsers.Count} 个用户");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建标准企业权限策略时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建机密文档权限策略
        /// </summary>
        /// <param name="allowedUsers">允许的用户列表</param>
        /// <param name="viewers">只读用户列表</param>
        /// <param name="expirationDays">过期天数</param>
        /// <returns>是否创建成功</returns>
        public bool CreateConfidentialDocumentPolicy(
            List<string> allowedUsers, 
            List<string> viewers, 
            int expirationDays = 7)
        {
            try
            {
                // 启用权限管理
                if (!EnablePermissionManagement())
                {
                    return false;
                }

                // 为允许的用户添加读写权限
                foreach (var user in allowedUsers)
                {
                    var permissions = new List<MsoPermission>
                    {
                        MsoPermission.msoPermissionRead,
                        MsoPermission.msoPermissionEdit,
                        MsoPermission.msoPermissionExtract
                    };

                    var expirationDate = DateTime.Now.AddDays(expirationDays);

                    AddUserPermission(user, permissions, expirationDate);
                }

                // 为只读用户添加读取权限
                foreach (var viewer in viewers)
                {
                    var permissions = new List<MsoPermission>
                    {
                        MsoPermission.msoPermissionRead
                    };

                    var expirationDate = DateTime.Now.AddDays(expirationDays);

                    AddUserPermission(viewer, permissions, expirationDate);
                }

                Console.WriteLine($"机密文档权限策略已创建");
                Console.WriteLine($"  读写用户: {allowedUsers.Count} 个");
                Console.WriteLine($"  只读用户: {viewers.Count} 个");
                Console.WriteLine($"  过期时间: {expirationDays} 天");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建机密文档权限策略时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取权限管理状态
        /// </summary>
        /// <returns>权限管理状态</returns>
        public PermissionManagementStatus GetPermissionManagementStatus()
        {
            var status = new PermissionManagementStatus();

            try
            {
                status.IsEnabled = _document.Permission.Enabled;
                status.UserCount = _document.Permission.Count;
                status.RequestPermissionUrl = _document.Permission.RequestPermissionURL;
                status.PermissionFromPolicy = _document.Permission.PermissionFromPolicy;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取权限管理状态时出错: {ex.Message}");
                status.ErrorMessage = ex.Message;
            }

            return status;
        }

        /// <summary>
        /// 创建权限报告
        /// </summary>
        /// <returns>权限报告</returns>
        public PermissionReport CreatePermissionReport()
        {
            var report = new PermissionReport();

            try
            {
                report.DocumentTitle = _document.Name;
                report.ReportDate = DateTime.Now;
                report.Status = GetPermissionManagementStatus();
                report.Users = GetAllUserPermissions();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建权限报告时出错: {ex.Message}");
                report.ErrorMessage = ex.Message;
            }

            return report;
        }
    }

    /// <summary>
    /// 用户权限信息类
    /// </summary>
    public class UserPermissionInfo
    {
        /// <summary>
        /// 用户ID
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// 权限值
        /// </summary>
        public int PermissionValue { get; set; }

        /// <summary>
        /// 过期日期
        /// </summary>
        public DateTime ExpirationDate { get; set; }

        /// <summary>
        /// 生成权限信息报告
        /// </summary>
        /// <returns>信息报告</returns>
        public string GenerateReport()
        {
            return $"用户权限信息:\n" +
                   $"  用户ID: {UserId}\n" +
                   $"  权限值: {PermissionValue}\n" +
                   $"  过期日期: {ExpirationDate:yyyy-MM-dd}";
        }
    }

    /// <summary>
    /// 权限管理状态类
    /// </summary>
    public class PermissionManagementStatus
    {
        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// 用户数量
        /// </summary>
        public int UserCount { get; set; }

        /// <summary>
        /// 请求权限URL
        /// </summary>
        public string RequestPermissionUrl { get; set; }

        /// <summary>
        /// 权限是否来自策略
        /// </summary>
        public bool PermissionFromPolicy { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成状态报告
        /// </summary>
        /// <returns>状态报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"权限管理状态检查失败: {ErrorMessage}";
            }

            return $"权限管理状态报告:\n" +
                   $"  是否启用: {IsEnabled}\n" +
                   $"  用户数量: {UserCount}\n" +
                   $"  请求权限URL: {RequestPermissionUrl}\n" +
                   $"  权限来自策略: {PermissionFromPolicy}";
        }
    }

    /// <summary>
    /// 权限报告类
    /// </summary>
    public class PermissionReport
    {
        /// <summary>
        /// 文档标题
        /// </summary>
        public string DocumentTitle { get; set; }

        /// <summary>
        /// 报告日期
        /// </summary>
        public DateTime ReportDate { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        public PermissionManagementStatus Status { get; set; }

        /// <summary>
        /// 用户列表
        /// </summary>
        public List<UserPermissionInfo> Users { get; set; } = new List<UserPermissionInfo>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成权限报告
        /// </summary>
        /// <returns>权限报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"生成权限报告失败: {ErrorMessage}";
            }

            var userReports = Users.Select(u => u.GenerateReport()).ToList();
            var usersReport = userReports.Any() ? string.Join("\n\n", userReports) : "无用户权限信息";

            return $"权限报告\n" +
                   $"文档标题: {DocumentTitle}\n" +
                   $"报告日期: {ReportDate:yyyy-MM-dd HH:mm:ss}\n\n" +
                   $"{Status.GenerateReport()}\n\n" +
                   $"用户权限信息:\n{usersReport}";
        }
    }
}
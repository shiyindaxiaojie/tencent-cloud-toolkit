import os
import logging
from typing import Dict, List
from dotenv import load_dotenv
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.clb.v20180317 import clb_client, models as clb_models
from tencentcloud.cvm.v20170312 import cvm_client, models as cvm_models
from tencentcloud.cfs.v20190719 import cfs_client, models as cfs_models
from tencentcloud.mariadb.v20170312 import mariadb_client, models as mariadb_models
from tencentcloud.redis.v20180412 import redis_client, models as redis_models
from tencentcloud.es.v20180416 import es_client, models as es_models
from tencentcloud.ckafka.v20190819 import ckafka_client, models as ckafka_models
from kubernetes import client as k8s_client, config as k8s_config

def setup_logging():
    log_dir = 'target'
    log_file = os.path.join(log_dir, 'ip-locator.log')

    try:
        os.makedirs(log_dir, exist_ok=True)
    except OSError as e:
        print(f"无法创建日志目录 {log_dir}: {str(e)}")
        log_dir = ''
        log_file = os.path.join(log_dir, 'ip-locator.log')

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


logger = setup_logging()


class TencentCloudIPLocator:
    def __init__(self):
        # 从 .env 加载腾讯云凭证
        load_dotenv()
        self.secret_id = os.getenv('TENCENTCLOUD_SECRET_ID')
        self.secret_key = os.getenv('TENCENTCLOUD_SECRET_KEY')
        self.region = os.getenv('TENCENTCLOUD_REGION', 'ap-guangzhou')

        # 加载 K8s 配置
        self.k8s_config_path = os.getenv('K8S_CONFIG_PATH', '~/.kube/config')

        if not all([self.secret_id, self.secret_key]):
            logger.error("腾讯云凭证未配置，请在 .env 文件中设置 TENCENTCLOUD_SECRET_ID 和 TENCENTCLOUD_SECRET_KEY")
            raise ValueError("Missing Tencent Cloud credentials")

        # 初始化腾讯云凭证
        self.cred = credential.Credential(self.secret_id, self.secret_key)

    def query_clb_by_ip(self, ip: str) -> List[Dict]:
        """查询 CLB 负载均衡"""
        try:
            client = clb_client.ClbClient(self.cred, self.region)

            req = clb_models.DescribeLoadBalancersRequest()
            req.LoadBalancerType = "OPEN"  # 公网类型
            req.LoadBalancerVips = [ip]

            resp = client.DescribeLoadBalancers(req)
            clb_instances = []

            for lb in resp.LoadBalancerSet:
                clb_instances.append({
                    "type": "CLB",
                    "instance_id": lb.LoadBalancerId,
                    "instance_name": lb.LoadBalancerName,
                    "vip": lb.LoadBalancerVips[0] if lb.LoadBalancerVips else None,
                    "status": lb.Status,
                    "region": lb.Zones
                })

            if not clb_instances:
                req = clb_models.DescribeLoadBalancersRequest()
                req.LoadBalancerType = "INTERNAL"  # 内网类型
                req.LoadBalancerVips = [ip]

                resp = client.DescribeLoadBalancers(req)
                for lb in resp.LoadBalancerSet:
                    clb_instances.append({
                        "type": "CLB",
                        "instance_id": lb.LoadBalancerId,
                        "instance_name": lb.LoadBalancerName,
                        "vip": lb.LoadBalancerVips[0] if lb.LoadBalancerVips else None,
                        "status": lb.Status,
                        "region": lb.Zones
                    })

            logger.info(f"CLB 匹配 IP {ip}，查询到 {len(clb_instances)} 个")
            return clb_instances

        except TencentCloudSDKException as e:
            logger.error(f"CLB 匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_cvm_by_ip(self, ip: str) -> List[Dict]:
        """查询 CVM 服务器"""
        try:
            client = cvm_client.CvmClient(self.cred, self.region)
            req = cvm_models.DescribeInstancesRequest()
            req.Filters = [{"Name": "private-ip-address", "Values": [ip]}]

            resp = client.DescribeInstances(req)
            instances = []
            for instance in resp.InstanceSet:
                instances.append({
                    "type": "CVM",
                    "instance_id": instance.InstanceId,
                    "instance_name": instance.InstanceName,
                    "private_ip": instance.PrivateIpAddresses[0] if instance.PrivateIpAddresses else None,
                    "public_ip": instance.PublicIpAddresses[0] if instance.PublicIpAddresses else None,
                    "region": instance.Placement.Zone,
                    "status": instance.InstanceState,
                    "create_time": instance.CreatedTime
                })

            if not instances:
                req = cvm_models.DescribeInstancesRequest()
                req.Filters = [{"Name": "public-ip-address", "Values": [ip]}]

                resp = client.DescribeInstances(req)
                instances = []
                for instance in resp.InstanceSet:
                    instances.append({
                        "type": "CVM",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "private_ip": instance.PrivateIpAddresses[0] if instance.PrivateIpAddresses else None,
                        "public_ip": instance.PublicIpAddresses[0] if instance.PublicIpAddresses else None,
                        "region": instance.Placement.Zone,
                        "status": instance.InstanceState,
                        "create_time": instance.CreatedTime
                    })

            logger.info(f"CVM 匹配 IP {ip}，查询到 {len(instances)} 个")
            return instances
        except TencentCloudSDKException as e:
            logger.error(f"CVM 匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_cfs_by_ip(self, ip: str) -> List[Dict]:
        """查询 CFS 文件系统"""
        try:
            client = cfs_client.CfsClient(self.cred, self.region)
            req = cfs_models.DescribeCfsFileSystemsRequest()
            resp = client.DescribeCfsFileSystems(req)

            matched_instances = []
            for fs in resp.FileSystems:
                client_req = cfs_models.DescribeCfsFileSystemClientsRequest()
                client_req.FileSystemId = fs.FileSystemId
                client_resp = client.DescribeCfsFileSystemClients(client_req)

                for client_info in client_resp.ClientList:
                    if client_info.ClientIp == ip or client_info.CfsVip == ip:
                        matched_instances.append({
                            "type": "CFS",
                            "instance_id": fs.FileSystemId,
                            "instance_name": fs.FsName,
                            "vip": client_info.CfsVip,
                            "client_ip": client_info.ClientIp,
                            "region": fs.Zone,
                            "status": fs.LifeCycleState
                        })

            logger.info(f"CFS 匹配 IP {ip}，查询到 {len(matched_instances)} 个")
            return matched_instances

        except TencentCloudSDKException as e:
            logger.error(f"CFS 匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_mariadb_by_ip(self, ip: str) -> List[Dict]:
        """查询 MariaDB 数据库"""
        try:
            client = mariadb_client.MariadbClient(self.cred, self.region)
            req = mariadb_models.DescribeDBInstancesRequest()
            resp = client.DescribeDBInstances(req)
            matched_instances = []
            for instance in resp.Instances:
                if instance.Vip == ip:
                    matched_instances.append({
                        "type": "MariaDB",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "vip": instance.Vip,
                        "port": instance.Vport,
                        "region": instance.Region,
                        "status": instance.Status
                    })
            logger.info(f"MariaDB 匹配 VIP {ip}，查询到 {len(matched_instances)} 个")
            return matched_instances
        except TencentCloudSDKException as e:
            logger.error(f"MariaDB 匹配 VIP {ip} 发生错误: {str(e)}")
            return []

    def query_redis_by_ip(self, ip: str) -> List[Dict]:
        """查询 Redis 数据库"""
        try:
            client = redis_client.RedisClient(self.cred, self.region)
            req = redis_models.DescribeInstancesRequest()
            resp = client.DescribeInstances(req)

            matched_instances = []
            for instance in resp.InstanceSet:
                if instance.WanIp == ip or instance.Vip6 == ip:
                    matched_instances.append({
                        "type": "Redis",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "vip": instance.WanIp,
                        "port": instance.Port,
                        "region": instance.Region,
                        "status": instance.Status
                    })

            logger.info(f"Redis 实例匹配 IP {ip}，查询到 {len(matched_instances)} 个")
            return matched_instances

        except TencentCloudSDKException as e:
            logger.error(f"Redis 实例匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_ckafka_by_ip(self, ip: str) -> List[Dict]:
        """查询 CKafka 消息队列"""
        try:
            client = ckafka_client.CkafkaClient(self.cred, self.region)
            req = ckafka_models.DescribeInstanceAttributesRequest()

            list_req = ckafka_models.DescribeInstancesRequest()
            list_resp = client.DescribeInstances(list_req)

            matched_instances = []
            for instance in list_resp.Result.InstanceList:
                req.InstanceId = instance.InstanceId
                resp = client.DescribeInstanceAttributes(req)

                if resp.Result.Vip == ip:
                    matched_instances.append({
                        "type": "CKafka",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "vip": resp.Result.Vip,
                        "port": resp.Result.Vport,
                        "region": "N/A",
                        "status": instance.Status
                    })

            logger.info(f"CKafka 匹配 IP {ip}，查询到 {len(matched_instances)} 个")
            return matched_instances

        except TencentCloudSDKException as e:
            logger.error(f"CKafka 匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_es_by_ip(self, ip: str) -> List[Dict]:
        """查询 Elasticsearch 搜索引擎"""
        try:
            client = es_client.EsClient(self.cred, self.region)
            req = es_models.DescribeInstancesRequest()
            resp = client.DescribeInstances(req)

            matched_instances = []
            for instance in resp.InstanceList:
                if instance.KibanaUrl and ip in instance.KibanaUrl:
                    matched_instances.append({
                        "type": "Elasticsearch",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "vip": instance.KibanaPrivateAccess,
                        "port": "N/A",
                        "region": instance.Zone,
                        "status": instance.Status
                    })
                elif instance.EsVip == ip:
                    matched_instances.append({
                        "type": "Elasticsearch",
                        "instance_id": instance.InstanceId,
                        "instance_name": instance.InstanceName,
                        "vip": instance.EsVip,
                        "port": instance.EsPort,
                        "region": instance.Zone,
                        "status": instance.Status
                    })

            logger.info(f"Elasticsearch 匹配 IP {ip}，查询到 {len(matched_instances)} 个")
            return matched_instances

        except TencentCloudSDKException as e:
            logger.error(f"Elasticsearch 匹配 IP {ip} 发生错误: {str(e)}")
            return []

    def query_all_resources(self, ip: str) -> Dict:
        """查询所有资源类型"""
        logger.info(f"开始查询 IP {ip} 绑定的资源信息")
        result = {
            "ip": ip,
            "clb": self.query_clb_by_ip(ip),
            "cvm": self.query_cvm_by_ip(ip),
            "mariadb": self.query_mariadb_by_ip(ip),
            "redis": self.query_redis_by_ip(ip),
            "k8s": self.query_k8s_pods_by_ip(ip)
        }
        return result

    def query_k8s_pods_by_ip(self, ip: str) -> List[Dict]:
        """遍历所有 K8s 上下文查询匹配 IP 的 Pod"""
        matched_pods = []
        try:
            # 加载 kubeconfig 并获取所有上下文
            contexts, _ = k8s_config.list_kube_config_contexts(config_file=os.path.expanduser(self.k8s_config_path))
            if not contexts:
                logger.warning("K8s 配置文件中未找到任何上下文")
                return matched_pods

            for ctx in contexts:
                ctx_name = ctx['name']
                try:
                    # 为每个上下文创建新配置
                    k8s_config.load_kube_config(context=ctx_name, config_file=os.path.expanduser(self.k8s_config_path))
                    v1 = k8s_client.CoreV1Api()

                    # 查询所有命名空间的 Pod
                    ret = v1.list_pod_for_all_namespaces(watch=False)
                    for pod in ret.items:
                        if pod.status.pod_ip == ip:
                            matched_pods.append({
                                "type": "EKS",
                                "cluster_id": ctx_name,
                                "namespace": pod.metadata.namespace,
                                "pod_name": pod.metadata.name,
                                "container_name": pod.spec.containers[0].name if pod.spec.containers else None,
                                "host_ip": pod.status.host_ip,
                                "pod_ip": pod.status.pod_ip,
                                "status": pod.status.phase
                            })
                except Exception as e:
                    logger.error(f"查询 K8s 上下文 {ctx_name} 时发生错误: {str(e)}")
                    continue

            logger.info(f"K8s 匹配 Pod IP {ip}，查询到 {len(matched_pods)} 个")
            return matched_pods
        except Exception as e:
            logger.error(f"K8s 匹配 Pod IP {ip} 发生全局错误: {str(e)}")
            return matched_pods

    def query_all_resources(self, ip: str) -> Dict:
        """查询所有资源类型"""
        logger.info(f"开始查询 IP {ip} 绑定的资源信息")
        result = {
            "ip": ip,
            "clb": self.query_clb_by_ip(ip),
            "cvm": self.query_cvm_by_ip(ip),
            "cfs": self.query_cfs_by_ip(ip),
            "mariadb": self.query_mariadb_by_ip(ip),
            "redis": self.query_redis_by_ip(ip),
            "ckafka": self.query_ckafka_by_ip(ip),
            "elasticsearch": self.query_es_by_ip(ip),
            "k8s": self.query_k8s_pods_by_ip(ip)
        }
        return result


if __name__ == "__main__":
    try:
        locator = TencentCloudIPLocator()
        while True:
            ip_to_query = input("\n请输入要查询的 IP 地址（或输入 q 退出）: ").strip()
            if ip_to_query.lower() == 'q':
                break

            if not all(part.isdigit() for part in ip_to_query.split('.')):
                print("错误：请输入有效的 IPv4 地址")
                continue

            result = locator.query_all_resources(ip_to_query)

            print("\n查询结果:")
            for resource_type in ["k8s", "clb", "cvm", "cfs", "mariadb", "redis", "ckafka", "elasticsearch"]:
                if result[resource_type]:
                    for item in result[resource_type]:
                        print(f"- 资源类型：{item['type']}")

                        # 腾讯云资源
                        if 'instance_id' in item:
                            print(f"  实例ID: {item.get('instance_id')}")
                            print(f"  实例名称: {item.get('instance_name', 'N/A')}")

                        if 'region' in item:
                            print(f"  区域: {item.get('region')}")

                        if 'private_ip' in item:
                            print(f"  内网IP: {item.get('private_ip', 'N/A')}")

                        if 'public_ip' in item:
                            print(f"  外网IP: {item.get('public_ip', 'N/A')}")

                        if 'vip' in item:
                            print(f"  VIP: {item.get('vip')}")

                        if 'port' in item:
                            print(f"  端口: {item.get('port')}")

                        if 'client_ip' in item:
                            print(f"  客户端IP: {item.get('client_ip')}")

                        # K8s 资源
                        if 'cluster_id' in item:
                            print(f"  集群ID: {item.get('cluster_id')}")

                        if 'namespace' in item:
                            print(f"  命名空间: {item.get('namespace')}")

                        if 'container_name' in item:
                            print(f"  容器名称: {item.get('container_name')}")

                        if 'pod_ip' in item:
                            print(f"  Pod IP: {item.get('pod_ip')}")

            break

    except KeyboardInterrupt:
        print("\n程序已退出")
    except Exception as e:
        logger.error(f"程序运行异常: {str(e)}")
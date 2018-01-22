#
# Author:: Paul Morton (<pmorton@biaprotect.com>)
# Copyright:: 2011-2017, Business Intelligence Associates, Inc.
# Copyright:: 2017-2018, Chef Software, Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

require "chef/resource"

class Chef
  class Resource
    # A resource to set applications to run at logon on Windows
    class WindowsAutorun < Chef::Resource
      resource_name :windows_auto_run
      provides :windows_auto_run

      property :program, String
      property :args, String
      property :root,
               Symbol,
               equal_to: %i(machine user),
               coerce: proc { |x| x.to_sym },
               default: :machine

      action :create do
        declare_resource(:registry_key, registry_path) do
          values [{
            name: new_resource.name,
            type: :string,
            data: "\"#{new_resource.program}\" #{new_resource.args}",
          }]
          action :create
        end
      end

      action :remove do
        declare_resource(:registry_key, registry_path) do
          values [{
            name: new_resource.name,
            type: :string,
            data: '',
          }]
          action :delete
        end
      end

      action_class do
        # determine the full registry path based on the root property
        # @return [String]
        def registry_path
          { machine: 'HKLM', user: 'HKCU' }[new_resource.root] + \
            '\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run'
        end
      end
    end
  end
end

# Author:: Stuart Preston (<stuart@chef.io>)
# Copyright:: Copyright 2018, Chef Software, Inc.
# License:: Apache License, Version 2.0
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

require "win32ole"
require "json"

class Chef
  module Mixin
    module PowershellExecute

      # Run a command under PowerShell via a managed (.NET) COM interop API.
      # This implementation requires the managed dll to be registered on the target machine.
      # Required: .NET Framework 4.0 or compatible, 64 bit platform.
      # 
      # Typical command used to install the interop assembly into the registry:
      # C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe Chef.PowerShell.dll /codebase
      #
      # @param script [String] script to run
      # @return [String] JSON containing output
      def powershell_execute(*command_args)
        script = command_args.first
        options = command_args.last.is_a?(Hash) ? command_args.last : nil

        ps = WIN32OLE.new('Chef.PowerShell')
        outcome = ps.ExecuteScript(script)

        JSON.parse(outcome)
      end
    end
  end
end
